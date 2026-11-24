"""SN → 物料号 本地小助手（只监听 127.0.0.1:8765）

查询顺序：① SQLite 缓存（以前查到过的）→ ② 导入的对照表 → ③ 实时查 SAP「出库查询」。

SAP 连接参数（地址 / client / 查询应用名）放在本地 config.json，不进版本库；
格式见 config.example.json。没有 config.json 时小助手照常运行，只是不查 SAP。

SAP 实时查询方式：
  启动时用本机 Chrome + ./browser_profile 打开 SAP Web Dynpro 出库查询页。
  未登录时在弹出的窗口里手动登录（本程序不接触密码）。
  登录后从页面取 contextid / secure-id / 控件 ID，之后每次查询直接复现
  Web Dynpro 的 POST（输入框 Change + 查询按钮 Press），只读。

接口：
  GET  /health        → {ok, sap: starting|need_login|ready|error, ...}
  GET  /sn/{sn}       → {sn, material, model, source: cache|table|sap}；查不到 404；需要登录 401
  POST /import        → 上传 Excel/CSV，自动识别 SN 列和物料号列
  GET  /              → rt_scan.html
"""
from __future__ import annotations

import concurrent.futures as cf
import contextlib
import csv
import html
import io
import json
import os
import queue
import re
import sqlite3
import threading
import time
import urllib.parse
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse

BASE = Path(__file__).resolve().parent
# 测试用：SN_HELPER_DATA 换数据目录，SN_HELPER_PORT 换端口，SN_HELPER_NO_SAP=1 不连 SAP
DATA_DIR = Path(os.environ.get("SN_HELPER_DATA") or BASE)
PORT = int(os.environ.get("SN_HELPER_PORT") or 8765)
NO_SAP = os.environ.get("SN_HELPER_NO_SAP") == "1"
DB_PATH = DATA_DIR / "sn_cache.sqlite3"
SAVE_DIR = DATA_DIR / "data" / "autosave"   # 清点进度的硬盘存档（每个清点一个文件）
PROFILE = Path(os.environ.get("SN_HELPER_PROFILE") or BASE / "browser_profile")
HTML_FILE = BASE / "rt_scan.html"

CONFIG_FILE = BASE / "config.json"


def _load_config() -> dict | None:
    """读取本地 SAP 配置。缺少或不完整时返回 None（小助手照常运行，只是不查 SAP）。"""
    try:
        c = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        assert all(c.get(k) for k in ("sap_base_url", "sap_client", "query_app"))
        return c
    except Exception:
        return None


CFG = _load_config()
QUERY_APP = (CFG or {}).get("query_app", "")
QUERY_BUTTON = (CFG or {}).get("query_button_text", "查询")
CKCX_URL = (f"{CFG['sap_base_url'].rstrip('/')}/sap/bc/webdynpro/sap/{QUERY_APP}"
            f"?sap-client={CFG['sap_client']}&sap-language={CFG.get('sap_language', 'ZH')}") if CFG else ""
QUERY_TIMEOUT = 15  # 秒
MATERIAL_RE = re.compile(r"^\d{9}$")
SN_OK = re.compile(r"^[A-Z0-9-]{4,32}$")  # 出库查询输入框最长 32 位


def log(*a):
    print(time.strftime("[%H:%M:%S]"), *a, flush=True)


def norm_sn(s: str) -> str:
    return re.sub(r"\s+", "", s or "").upper()


# ───────────────────────── SQLite ─────────────────────────
_db_lock = threading.Lock()
_db = sqlite3.connect(DB_PATH, check_same_thread=False)
_db.executescript("""
CREATE TABLE IF NOT EXISTS cache(
  sn TEXT PRIMARY KEY, material TEXT NOT NULL, model TEXT, updated REAL);
CREATE TABLE IF NOT EXISTS imported(
  sn TEXT PRIMARY KEY, material TEXT NOT NULL, model TEXT, file TEXT, updated REAL);
""")
_db.commit()


def db_get(table: str, sn: str):
    with _db_lock:
        r = _db.execute(f"SELECT material, model FROM {table} WHERE sn=?", (sn,)).fetchone()
    return r


def db_put(table: str, rows: list[tuple], extra_file: str | None = None):
    now = time.time()
    with _db_lock:
        if table == "cache":
            _db.executemany("INSERT OR REPLACE INTO cache VALUES(?,?,?,?)",
                            [(s, m, md, now) for s, m, md in rows])
        else:
            _db.executemany("INSERT OR REPLACE INTO imported VALUES(?,?,?,?,?)",
                            [(s, m, md, extra_file, now) for s, m, md in rows])
        _db.commit()


# ───────────────────────── SAP 工作线程 ─────────────────────────
class NeedLogin(Exception):
    pass


class SapWorker(threading.Thread):
    """独占一个 Playwright 浏览器；所有 SAP 查询在这个线程里串行执行（一次一个）。"""

    def __init__(self):
        super().__init__(daemon=True, name="sap-worker")
        self.jobs: queue.Queue = queue.Queue()
        self.state = "starting"  # starting | need_login | ready | error
        self.detail = ""
        self.sess = None  # dict(url, secure_id, inp, btn)

    # —— 对外：提交查询，阻塞等待结果 ——
    def lookup(self, sn: str):
        fut: cf.Future = cf.Future()
        self.jobs.put((sn, fut))
        # 单次 SAP 请求本身限时 QUERY_TIMEOUT；这里的等待包含排队时间，所以放宽
        return fut.result(timeout=180)

    # —— 线程主体 ——
    def run(self):
        try:
            from playwright.sync_api import sync_playwright
            with sync_playwright() as p:
                self.pw = p
                self.ctx = None
                self._launch()
                self._open_app()
                self._loop()
        except Exception as e:
            self.state, self.detail = "error", str(e)
            log("SAP 工作线程出错（请关掉黑色窗口后重新双击 start.bat）：", e)
            while True:  # 把排队的请求都打回去
                sn, fut = self.jobs.get()
                fut.set_exception(RuntimeError(self.detail))

    # —— 浏览器生命周期：被用户关掉后自动重开 ——
    def _launch(self):
        last = None
        for attempt in range(4):  # 刚关掉的 Chrome 可能还占着 profile，稍等重试
            for ch in ("chrome", "msedge"):
                try:
                    ctx = self.pw.chromium.launch_persistent_context(
                        str(PROFILE), channel=ch, headless=False, viewport=None)
                    self.ctx, self.ctx_closed = ctx, False
                    ctx.on("close", lambda *_: setattr(self, "ctx_closed", True))
                    self.page = ctx.pages[0] if ctx.pages else ctx.new_page()
                    self.last_launch = time.time()
                    return
                except Exception as e:  # 没装这个浏览器 / profile 被占用
                    last = e
                    log(f"用 {ch} 启动浏览器失败：{e}".splitlines()[0])
            time.sleep(2)
        raise RuntimeError(f"无法启动 Chrome/Edge：{last}")

    def _alive(self) -> bool:
        """浏览器还在吗？标签页被关了就补开一个；整个窗口被关了返回 False。
        不在页面里执行 JS（页面跳转中会误报），而是靠浏览器的 close 事件判断。"""
        try:
            if not self.page.is_closed():
                self.page.wait_for_timeout(50)  # 让 Playwright 处理积压事件（关闭事件在这里送达）
            if self.ctx_closed:
                return False
            if self.page.is_closed():
                self.page = self.ctx.new_page()
            return not self.ctx_closed
        except Exception:
            return False

    def _recover(self):
        """浏览器被关掉后重开（会话随窗口一起没了，所以通常需要重新登录）。"""
        log("SAP 浏览器窗口被关掉了，正在重新打开…（请在新窗口里重新登录）")
        self.state, self.sess = "starting", None
        try:
            self.ctx.close()
        except Exception:
            pass
        self._launch()
        self._open_app()

    def _loop(self):
        while True:
            try:
                sn, fut = self.jobs.get(timeout=2)
            except queue.Empty:
                # 空闲时：先确认浏览器还在，再看是否登录好了。任何一步出错都只记日志，
                # 绝不能让异常冒出去把整个工作线程弄死（实测：窗口被关后曾因此卡死）。
                try:
                    if not self._alive():
                        if time.time() - getattr(self, "last_launch", 0) < 20:
                            continue  # 刚重开过又没了：别疯狂弹窗，等一会儿
                        self._recover()
                        continue
                    if self.state == "need_login" and self._try_capture_session(quiet=True):
                        # 刚登录完页面可能还在跳转，取到的参数不稳定（实测首查会超时）；
                        # 重新打开一次出库查询页再取
                        self._open_app()
                except Exception as e:
                    log("后台检查出错（会自动重试）：", str(e).splitlines()[0])
                continue
            try:
                if not self._alive():
                    self._recover()
                fut.set_result(self._query(sn))
            except Exception as e:
                fut.set_exception(e)

    def _ensure_page(self):
        try:
            if self.page.is_closed():
                raise RuntimeError
        except Exception:
            self.page = self.ctx.new_page()

    def _open_app(self):
        self._ensure_page()
        try:
            self.page.goto(CKCX_URL, timeout=30000)
            self.page.wait_for_timeout(1500)
        except Exception as e:
            log("打开出库查询页失败：", e)
        self._try_capture_session()

    def _try_capture_session(self, quiet=False) -> bool:
        """当前页面若已是登录后的出库查询页，就取出会话参数。"""
        try:
            self._ensure_page()
            pg = self.page
            if QUERY_APP not in pg.url:
                # 登录完成后可能停在别处，带回出库查询
                if not pg.query_selector("input[type=password]"):
                    pg.goto(CKCX_URL, timeout=30000)
                    pg.wait_for_timeout(1500)
            info = pg.evaluate("""(btnText) => {
              const pw = !!document.querySelector('input[type=password]');
              const f = document.forms[0];
              const sec = document.querySelector('input[name="sap-wd-secure-id"]');
              const inp = document.querySelector('input[ct="I"]');
              const btn = [...document.querySelectorAll('[ct="B"]')].find(b => b.innerText.trim() === btnText);
              return {pw, action: f && f.action, sec: sec && sec.value,
                      inp: inp && inp.id, btn: btn && btn.id, title: document.title};
            }""", QUERY_BUTTON)
        except Exception as e:
            if not quiet:
                log("读取页面失败：", e)
            info = {"pw": False}
        if info.get("pw") or not (info.get("sec") and info.get("inp") and info.get("btn")
                                  and QUERY_APP in (info.get("action") or "")):
            if self.state != "need_login":
                self.state = "need_login"
                self.sess = None
                log("⚠ 需要登录 SAP：请在弹出的浏览器窗口里手动登录（登录后会自动继续）。")
            return False
        self.sess = {"url": info["action"], "secure_id": info["sec"],
                     "inp": info["inp"], "btn": info["btn"]}
        if self.state != "ready":
            log(f"✔ SAP 已就绪（出库查询，控件 {info['inp']}/{info['btn']}）")
        self.state = "ready"
        return True

    def _post(self, sn: str):
        s = self.sess
        q = (f"InputField_Change~E002Id~E004{s['inp']}~E005Value~E004{sn}~E003~E002ResponseData~E004delta"
             "~E005EnqueueCardinality~E004single~E005Delay~E004full~E003~E002~E003~E001"
             f"Button_Press~E002Id~E004{s['btn']}~E003~E002ResponseData~E004delta~E005ClientAction~E004submit"
             "~E003~E002~E003")
        body = urllib.parse.urlencode({"sap-charset": "utf-8", "sap-wd-secure-id": s["secure_id"],
                                       "_stateful_": "X", "SAPEVENTQUEUE": q})
        return self.ctx.request.post(s["url"], data=body, timeout=QUERY_TIMEOUT * 1000, headers={
            "content-type": "application/x-www-form-urlencoded",
            "x-requested-with": "XMLHttpRequest"}, max_redirects=0)

    def _query(self, sn: str):
        if self.state != "ready" and not self._try_capture_session(quiet=True):
            raise NeedLogin()
        for attempt in (1, 2):
            t0 = time.time()
            try:
                r = self._post(sn)
                text = r.text() if r.status == 200 else ""
            except Exception as e:
                log(f"查询 {sn} 请求异常：{e}".splitlines()[0])
                r, text = None, ""
            ok = r is not None and r.status == 200 and f"'{self.sess['inp']}'" in text
            if ok:
                res = parse_ckcx(text, sn)
                log(f"SAP {sn} → {res['material'] if res else '查无'}  ({time.time() - t0:.1f}s)")
                return res
            # 会话可能过期 / 浏览器被关：重新打开页面取参数，再试一次
            log(f"查询 {sn} 返回异常（HTTP {r.status if r else '-'}），重新建立会话…")
            self.state = "starting"
            if self._alive():
                self._open_app()
            else:
                self._recover()
            if self.state != "ready":
                raise NeedLogin()
        raise RuntimeError("SAP 查询失败")


def _unesc(v: str) -> str:
    # SAP 把反斜杠也写成 HTML 实体（&#x5c;x2d），所以先 HTML 反转义，再解 \xNN
    v = html.unescape(v)
    v = re.sub(r"\\x([0-9a-fA-F]{2})", lambda m: chr(int(m.group(1), 16)), v)
    v = re.sub(r"\\u([0-9a-fA-F]{4})", lambda m: chr(int(m.group(1), 16)), v)
    return html.unescape(v).strip()


def parse_ckcx(text: str, sn: str):
    """主表每行的单元格顺序：物料, 序列号, 产品型号, 客户型号, …
    取「物料(9位数字) 紧接着 序列号==查询SN」的第一行；组件表顺序是 序列号, 物料，不会误匹配。"""
    vals = [_unesc(v) for v in re.findall(r"lsdata=\"\{0:'\w+',2:'([^']*)'", text)]
    hits = []
    for i in range(1, len(vals)):
        if norm_sn(vals[i]) == sn and MATERIAL_RE.match(vals[i - 1]):
            model = vals[i + 1] if i + 1 < len(vals) else ""
            hits.append((vals[i - 1], model))
    if not hits:
        return None
    material, model = hits[0]
    return {"material": material, "model": model,
            "others": sorted({m for m, _ in hits[1:] if m != material})}


# ───────────────────────── 导入对照表 ─────────────────────────
SN_HDR = re.compile(r"(序列号|sn|serial|s/n)", re.I)
MAT_HDR = re.compile(r"(物料|material|matnr|料号)", re.I)
MODEL_HDR = re.compile(r"(型号|model)", re.I)


def read_table(name: str, data: bytes) -> list[list[str]]:
    if name.lower().endswith((".xlsx", ".xlsm")):
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        rows = []
        for ws in wb.worksheets:
            for r in ws.iter_rows(values_only=True):
                rows.append(["" if c is None else str(c).strip() for c in r])
        return rows
    for enc in ("utf-8-sig", "gb18030", "latin-1"):
        try:
            txt = data.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    dialect = csv.Sniffer().sniff(txt[:4096], delimiters=",;\t") if txt.strip() else csv.excel
    return [[c.strip() for c in r] for r in csv.reader(io.StringIO(txt), dialect)]


def _clean(v: str) -> str:
    """去掉本工具导出 CSV 里防 Excel 变形用的 ="…" 包装。"""
    v = (v or "").strip()
    m = re.fullmatch(r'="(.*)"', v)
    return m.group(1).strip() if m else v


def _fmt_mat(v: str) -> str:
    v = _clean(v)
    return v[:-2] if v.endswith(".0") else v


def detect_columns(rows: list[list[str]]):
    """先看表头，再按内容统计：物料列 = 9 位数字最多的列；SN 列 = 字母数字混合、6–30 位最多的列。"""
    width = max((len(r) for r in rows), default=0)
    hdr_idx, sn_c, mat_c, model_c = None, None, None, None
    for ri, r in enumerate(rows[:10]):
        s = [i for i, c in enumerate(r) if SN_HDR.search(c) and len(c) < 20]
        m = [i for i, c in enumerate(r) if MAT_HDR.search(c) and len(c) < 20]
        if s and m:
            hdr_idx, sn_c, mat_c = ri, s[0], m[0]
            md = [i for i, c in enumerate(r) if MODEL_HDR.search(c) and len(c) < 20]
            model_c = md[0] if md else None
            break
    body = rows[hdr_idx + 1:] if hdr_idx is not None else rows
    if sn_c is None or mat_c is None:
        mat_score = [sum(bool(MATERIAL_RE.match(_fmt_mat(r[c]))) for r in body if c < len(r))
                     for c in range(width)]
        sn_score = [sum(bool(re.fullmatch(r"(?=.*[A-Za-z])(?=.*\d)[A-Za-z0-9\-]{6,30}", r[c]))
                        for r in body if c < len(r)) for c in range(width)]
        mat_c = max(range(width), key=lambda c: mat_score[c]) if width else None
        cand = [c for c in range(width) if c != mat_c]
        sn_c = max(cand, key=lambda c: sn_score[c]) if cand else None
        if mat_c is None or sn_c is None or not mat_score[mat_c] or not sn_score[sn_c]:
            raise ValueError("没能识别出 SN 列和物料号列，请给表头加上「序列号」「物料」")
    return body, sn_c, mat_c, model_c


# ───────────────────────── FastAPI ─────────────────────────
worker = SapWorker()


@contextlib.asynccontextmanager
async def lifespan(_app):
    if NO_SAP:
        worker.state, worker.detail = "error", "测试模式：未连接 SAP"
    elif not CFG:
        worker.state, worker.detail = "error", "未配置 SAP：把 config.example.json 复制为 config.json 并填写"
        log("⚠ 没有找到有效的 config.json，SAP 实时查询已关闭（缓存、对照表、存档照常可用）。")
    else:
        worker.start()
    SAVE_DIR.mkdir(parents=True, exist_ok=True)
    log(f"小助手已启动：http://127.0.0.1:{PORT}/   清点存档目录：{SAVE_DIR}")
    yield


app = FastAPI(title="SN helper", lifespan=lifespan, docs_url=None, redoc_url=None, openapi_url=None)

# ───────────────────────── 访问保护 ─────────────────────────
# 小助手只给本机的清点页面（同源）用。防三类攻击：
#  · DNS 重绑定：恶意域名解析到 127.0.0.1 → 校验 Host 头；
#  · 别的网站跨站提交（表单 / sendBeacon / fetch no-cors）→ 带 Origin 的请求必须是本页面；
#  · 别的网站用 <img>/no-cors GET 触发 SAP 查询 → API 必须带自定义头 X-SN-Helper（跨站无法添加）。
# 不开 CORS：其它来源（包括直接双击打开的 file:// 页面）一律读不到返回内容。
ALLOWED_HOSTS = {f"127.0.0.1:{PORT}", f"localhost:{PORT}"}
ALLOWED_ORIGINS = {f"http://{h}" for h in ALLOWED_HOSTS}


@app.middleware("http")
async def guard(request: Request, call_next):
    if request.headers.get("host", "") not in ALLOWED_HOSTS:
        return JSONResponse({"error": "拒绝：Host 不对"}, 403)
    origin = request.headers.get("origin")
    if origin is not None and origin not in ALLOWED_ORIGINS:
        return JSONResponse({"error": "拒绝：只允许本机清点页面访问"}, 403)
    if request.url.path != "/" and request.headers.get("x-sn-helper") != "1":
        # 页面关闭时的 sendBeacon 不能带自定义头：/state 凭同源 Origin 放行
        if not (request.url.path == "/state" and request.method == "POST" and origin in ALLOWED_ORIGINS):
            return JSONResponse({"error": "拒绝：缺少 X-SN-Helper 头"}, 403)
    return await call_next(request)


@app.get("/health")
def health():
    with _db_lock:
        nc = _db.execute("SELECT COUNT(*) FROM cache").fetchone()[0]
        ni = _db.execute("SELECT COUNT(*) FROM imported").fetchone()[0]
    return {"ok": True, "sap": worker.state, "detail": worker.detail,
            "cache": nc, "imported": ni}


@app.get("/sn/{raw_sn}")
def lookup(raw_sn: str, refresh: bool = False):
    sn = norm_sn(raw_sn)
    if not SN_OK.match(sn):  # 只把像 SN 的东西发给 SAP（也防止特殊字符混进 SAP 请求）
        return JSONResponse({"sn": sn, "error": "格式不像 SN（只允许字母、数字、横杠，4–32 位）"}, 400)
    if not refresh:
        r = db_get("cache", sn)
        if r:
            return {"sn": sn, "material": r[0], "model": r[1], "source": "cache"}
        r = db_get("imported", sn)
        if r:
            return {"sn": sn, "material": r[0], "model": r[1], "source": "table"}
    try:
        if NO_SAP or not CFG:
            raise RuntimeError(worker.detail or "未连接 SAP")
        res = worker.lookup(sn)
    except NeedLogin:
        return JSONResponse({"sn": sn, "error": "需要登录 SAP（在弹出的浏览器窗口里登录）"}, 401)
    except cf.TimeoutError:
        return JSONResponse({"sn": sn, "error": "SAP 查询超时"}, 504)
    except Exception as e:
        return JSONResponse({"sn": sn, "error": f"SAP 查询失败：{e}"}, 502)
    if not res:
        return JSONResponse({"sn": sn, "error": "SAP 出库查询无此 SN"}, 404)
    db_put("cache", [(sn, res["material"], res["model"])])
    out = {"sn": sn, "material": res["material"], "model": res["model"], "source": "sap"}
    if res["others"]:
        out["others"] = res["others"]
    return out


@app.post("/import")
async def import_table(file: UploadFile = File(...)):
    data = await file.read()
    try:
        rows = read_table(file.filename or "", data)
        body, sn_c, mat_c, model_c = detect_columns(rows)
    except Exception as e:
        raise HTTPException(400, f"无法读取：{e}")
    out, skipped = [], 0
    for r in body:
        sn = norm_sn(_clean(r[sn_c])) if sn_c < len(r) else ""
        mat = _fmt_mat(r[mat_c]) if mat_c < len(r) else ""
        # SN 必须像 SN（防止把汇总行、时间之类误当 SN 导进去），物料号必须是 9 位数字
        if not SN_OK.match(sn) or not MATERIAL_RE.match(mat):
            skipped += 1
            continue
        model = _clean(r[model_c]) if model_c is not None and model_c < len(r) else ""
        out.append((sn, mat, model))
    db_put("imported", out, extra_file=file.filename)
    log(f"导入 {file.filename}：{len(out)} 条，跳过 {skipped} 行")
    return {"imported": len(out), "skipped": skipped,
            "columns": {"sn": sn_c, "material": mat_c, "model": model_c}}


@app.get("/")
def index():
    return FileResponse(HTML_FILE, media_type="text/html; charset=utf-8")


# ───────────────────────── 清点进度硬盘存档 ─────────────────────────
# 页面每次改动后把整个清点状态 POST 过来，按清点 ID 各存一个 JSON 文件。
# 这样即使浏览器数据被清掉、换了浏览器，进度也能从这里找回。
SID_RE = re.compile(r"^[A-Za-z0-9_-]{1,64}$")
PREV_EVERY = 600  # 每个清点额外保留一份 ≥10 分钟前的旧版本（.prev），防止一次写坏
_save_lock = threading.Lock()


def _save_path(sid: str, prev: bool = False) -> Path:
    if not SID_RE.match(sid or ""):
        raise HTTPException(400, "非法的清点 ID")
    return SAVE_DIR / f"{sid}{'.prev' if prev else ''}.json"


@app.post("/state")
async def save_state(request: Request):
    raw = await request.body()  # 不看 content-type：页面关闭时用 sendBeacon(text/plain) 发
    if len(raw) > 30_000_000:
        raise HTTPException(413, "存档太大")
    try:
        obj = json.loads(raw)
        sid, state = obj["sid"], obj["state"]
        assert isinstance(state, dict) and isinstance(state.get("recs", []), list)
    except Exception:
        raise HTTPException(400, "存档格式不对")
    f, fp = _save_path(sid), _save_path(sid, prev=True)
    rec = {"sid": sid, "saved_at": obj.get("saved_at"), "server_time": time.time(), "state": state}
    with _save_lock:
        SAVE_DIR.mkdir(parents=True, exist_ok=True)
        if f.exists() and (not fp.exists() or time.time() - fp.stat().st_mtime > PREV_EVERY):
            os.replace(f, fp)
        tmp = f.with_suffix(".tmp")
        tmp.write_text(json.dumps(rec, ensure_ascii=False), encoding="utf-8")
        os.replace(tmp, f)  # 原子替换：写到一半断电也不会得到半个文件
    return {"ok": True, "sid": sid}


@app.get("/saves")
def list_saves(limit: int = 100):
    out = []
    files = sorted((p for p in SAVE_DIR.glob("*.json") if not p.name.endswith(".prev.json")),
                   key=lambda p: p.stat().st_mtime, reverse=True)[:limit]
    for p in files:
        try:
            st = json.loads(p.read_text(encoding="utf-8"))["state"]
        except Exception:
            continue
        out.append({"sid": p.stem, "mtime": p.stat().st_mtime, "rt": st.get("rt", ""),
                    "blind": bool(st.get("blind")), "items": len(st.get("items", [])),
                    "recs": len(st.get("recs", [])), "pend": len(st.get("pend", []))})
    return out


@app.get("/saves/{sid}")
def get_save(sid: str, prev: bool = False):
    f = _save_path(sid, prev)
    if not f.exists():
        raise HTTPException(404, "没有这个存档")
    return JSONResponse(json.loads(f.read_text(encoding="utf-8")))


PAGE_URL = f"http://127.0.0.1:{PORT}/"


def _already_running() -> bool:
    """端口上是否已经有小助手（或别的程序）在跑。"""
    import socket
    with socket.socket() as s:
        s.settimeout(0.5)
        return s.connect_ex(("127.0.0.1", PORT)) == 0


if __name__ == "__main__":
    import sys
    import webbrowser

    if _already_running():
        log("小助手已经在运行了（可能是之前打开的窗口），直接打开清点页面。")
        log("如果页面打不开，请关掉其它 SN 小助手窗口后再双击 start.bat。")
        webbrowser.open(PAGE_URL)
        sys.exit(0)
    print("=" * 60)
    print(" SN 小助手 —— 这个窗口别关，关掉即停止小助手")
    print(" 如果弹出的 Chrome 窗口停在 SAP 登录页，请在那里手动登录")
    print("=" * 60, flush=True)
    if not NO_SAP:
        threading.Timer(2.0, webbrowser.open, args=(PAGE_URL,)).start()
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=PORT, log_level="warning")
