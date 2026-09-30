#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
电商融 · 拼多多订单导出 —— 服务器版
=====================================
保留原有架构：Flask 网页 -> pdd_export.py -> Playwright。

新增：
  - 每个浏览器独立保存自己的电商融账号密码（服务端加密）
  - 多人任务队列，显示排队位置、正在执行和当前步骤
  - 每用户独立浏览器 profile / CDP 端口，避免串账号
  - 报表完成后通过浏览器下载到同事自己的电脑

运行：
  python webui.py
环境变量：
  PDD_WEB_ACCESS_PASSWORD  可选的站点访问密码；设置后打开网页需先输入
  PDD_SECRET_KEY           服务端加密主密钥；不设置则自动生成 data/secret.key
  PDD_WEB_HOST             监听地址，默认 0.0.0.0
  PDD_WEB_PORT             监听端口，默认 8765
  PDD_OPEN_BROWSER         默认 1，服务器可设为 0
"""

import base64
import hashlib
import hmac
import io
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import threading
import time
import uuid
import webbrowser
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

from cryptography.fernet import Fernet, InvalidToken
from flask import Flask, Response, jsonify, request, send_file, session

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("PDD_WEB_DATA_DIR") or (BASE_DIR / "data")).resolve()
JOBS_DIR = DATA_DIR / "jobs"
PROFILES_DIR = DATA_DIR / "profiles"
# 任务完成后报表归档目录（对外整理用；网页下载仍走任务目录，删除归档不影响下载）
ORDER_ARCHIVE = Path(os.environ.get("PDD_ORDER_ARCHIVE") or r"E:\order")
DB_PATH = DATA_DIR / "jobs.db"
KEY_PATH = DATA_DIR / "secret.key"
APP_SECRET_PATH = DATA_DIR / "flask_secret.key"
MAX_BACKOFF_DAYS = 92
CLIENT_COOKIE = "pdd_client_id"
FINISHED_STATUSES = ("已完成", "部分失败", "失败", "已取消")
ACCESS_PASSWORD = os.environ.get("PDD_WEB_ACCESS_PASSWORD", "")
LINK_TOKEN = os.environ.get("PDD_WEB_LINK_TOKEN", "").strip()
LINK_TOKEN_COOKIE = "pdd_link_token"
WORKER_WAKE = threading.Event()
DB_LOCK = threading.RLock()
RUNNING_LOCK = threading.RLock()
RUNNING = {"job_id": "", "proc": None}
WORKER_STARTED = False
WORKER_LOCK = threading.Lock()

for d in (DATA_DIR, JOBS_DIR, PROFILES_DIR):
    d.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)


# ---------------------------------------------------------------------------
# 数据库与加密
# ---------------------------------------------------------------------------

def _secret_bytes() -> bytes:
    env_key = os.environ.get("PDD_SECRET_KEY", "").strip()
    if env_key:
        return hashlib.sha256(env_key.encode("utf-8")).digest()
    if KEY_PATH.exists():
        raw = KEY_PATH.read_bytes().strip()
        if len(raw) == 44:
            try:
                return base64.urlsafe_b64decode(raw)
            except Exception:
                pass
    key = Fernet.generate_key()
    KEY_PATH.write_bytes(key)
    try:
        os.chmod(KEY_PATH, 0o600)
    except Exception:
        pass
    return base64.urlsafe_b64decode(key)


_FERNET = Fernet(base64.urlsafe_b64encode(_secret_bytes()))


def _encrypt_text(value: str) -> str:
    return _FERNET.encrypt(value.encode("utf-8")).decode("ascii")


def _decrypt_text(value: str) -> str:
    if not value:
        return ""
    try:
        return _FERNET.decrypt(value.encode("ascii")).decode("utf-8")
    except (InvalidToken, ValueError, TypeError):
        return ""


def _flask_secret() -> bytes:
    env_key = os.environ.get("PDD_FLASK_SECRET", "").strip()
    if env_key:
        return env_key.encode("utf-8")
    if APP_SECRET_PATH.exists():
        return APP_SECRET_PATH.read_bytes().strip()
    secret = base64.urlsafe_b64encode(os.urandom(32))
    APP_SECRET_PATH.write_bytes(secret)
    try:
        os.chmod(APP_SECRET_PATH, 0o600)
    except Exception:
        pass
    return secret


app.secret_key = _flask_secret()


def _db():
    con = sqlite3.connect(DB_PATH, timeout=30, check_same_thread=False)
    con.row_factory = sqlite3.Row
    return con


def _init_db() -> None:
    with DB_LOCK, _db() as con:
        con.executescript(
            """
            CREATE TABLE IF NOT EXISTS jobs (
                seq INTEGER PRIMARY KEY AUTOINCREMENT,
                id TEXT NOT NULL UNIQUE,
                client_id TEXT NOT NULL,
                submitter TEXT NOT NULL,
                created_at TEXT NOT NULL,
                started_at TEXT,
                finished_at TEXT,
                status TEXT NOT NULL,
                progress TEXT NOT NULL DEFAULT '',
                shops TEXT NOT NULL,
                account TEXT NOT NULL DEFAULT '',
                password_enc TEXT NOT NULL DEFAULT '',
                start_date TEXT NOT NULL,
                end_date TEXT NOT NULL,
                log_path TEXT NOT NULL DEFAULT '',
                out_dir TEXT NOT NULL DEFAULT '',
                results TEXT NOT NULL DEFAULT '[]',
                error TEXT NOT NULL DEFAULT ''
            );
            CREATE INDEX IF NOT EXISTS idx_jobs_status_seq ON jobs(status, seq);
            CREATE INDEX IF NOT EXISTS idx_jobs_client_seq ON jobs(client_id, seq);

            CREATE TABLE IF NOT EXISTS credentials (
                client_id TEXT PRIMARY KEY,
                submitter TEXT NOT NULL DEFAULT '',
                account TEXT NOT NULL DEFAULT '',
                password_enc TEXT NOT NULL DEFAULT '',
                updated_at TEXT NOT NULL
            );
            """
        )
        cols = {r["name"] for r in con.execute("PRAGMA table_info(jobs)").fetchall()}
        if "account" not in cols:
            con.execute("ALTER TABLE jobs ADD COLUMN account TEXT NOT NULL DEFAULT ''")
        if "password_enc" not in cols:
            con.execute("ALTER TABLE jobs ADD COLUMN password_enc TEXT NOT NULL DEFAULT ''")
        con.execute(
            """UPDATE jobs SET password_enc=''
               WHERE status IN ('已完成','部分失败','失败','已取消')"""
        )
        con.execute(
            """UPDATE jobs SET status='失败', finished_at=?,
                      progress='服务重启，任务已中断',
                      error='服务重启，任务已中断'
               WHERE status='正在执行'""",
            (_now(),),
        )


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _row_to_job(row: sqlite3.Row, client_id: str, queue_positions: dict) -> dict:
    try:
        raw_results = json.loads(row["results"] or "[]")
    except Exception:
        raw_results = []
    results = []
    for idx, item in enumerate(raw_results):
        status = item.get("status", "")
        result = {
            "shop": item.get("shop", ""),
            "status": status,
            "orders": item.get("orders", ""),
            "file": item.get("file", ""),
            "error": item.get("error", ""),
            "download_url": "",
        }
        if item.get("file_rel") and row["client_id"] == client_id and status == "成功":
            result["download_url"] = f"/api/jobs/{row['id']}/files/{idx}"
        results.append(result)
    shops = []
    try:
        shops = json.loads(row["shops"] or "[]")
    except Exception:
        pass
    return {
        "id": row["id"],
        "submitter": row["submitter"],
        "created_at": row["created_at"],
        "started_at": row["started_at"] or "",
        "finished_at": row["finished_at"] or "",
        "status": row["status"],
        "progress": row["progress"] or "",
        "queue_position": queue_positions.get(row["id"], 0),
        "shops": shops,
        "start_date": row["start_date"],
        "end_date": row["end_date"],
        "results": results,
        "error": row["error"] or "",
        "is_owner": row["client_id"] == client_id,
    }


def _jobs_for_client(client_id: str, limit: int = 100) -> list[dict]:
    # 已结束的任务只在列表保留 1 小时（足够点下载），之后不再显示；
    # 24 小时后由 _purge_old_jobs 连内部文件一起清掉（E:\order 归档永久保留）。
    hide_cutoff = (datetime.now() - timedelta(hours=1)).strftime("%Y-%m-%d %H:%M:%S")
    with DB_LOCK, _db() as con:
        queued = con.execute(
            "SELECT id FROM jobs WHERE status='排队中' ORDER BY seq"
        ).fetchall()
        queue_positions = {r["id"]: i + 1 for i, r in enumerate(queued)}
        rows = con.execute(
            "SELECT * FROM jobs ORDER BY seq DESC LIMIT ?", (limit,)
        ).fetchall()
    out = []
    for row in rows:
        if (row["status"] in FINISHED_STATUSES and row["finished_at"]
                and row["finished_at"] < hide_cutoff):
            continue
        out.append(_row_to_job(row, client_id, queue_positions))
    return out


def _get_job(job_id: str) -> sqlite3.Row | None:
    with DB_LOCK, _db() as con:
        return con.execute("SELECT * FROM jobs WHERE id=?", (job_id,)).fetchone()


def _profile_for(client_id: str) -> dict:
    with DB_LOCK, _db() as con:
        row = con.execute(
            "SELECT * FROM credentials WHERE client_id=?", (client_id,)
        ).fetchone()
    if not row:
        return {"submitter": "", "account": "", "has_password": False}
    return {
        "submitter": row["submitter"] or "",
        "account": row["account"] or "",
        "has_password": bool(row["password_enc"]),
    }


def _saved_credentials(client_id: str) -> tuple[str, str, str]:
    with DB_LOCK, _db() as con:
        row = con.execute(
            "SELECT * FROM credentials WHERE client_id=?", (client_id,)
        ).fetchone()
    if not row:
        return "", "", ""
    return (
        row["submitter"] or "",
        row["account"] or "",
        _decrypt_text(row["password_enc"] or ""),
    )


def _upsert_credentials(client_id: str, submitter: str,
                        account: str, password: str | None,
                        remember_password: bool) -> None:
    """姓名与账号始终保存（下次免填）；密码仅在 remember_password 时保存/保留。"""
    new_enc = _encrypt_text(password) if password else ""
    with DB_LOCK, _db() as con:
        old = con.execute(
            "SELECT password_enc FROM credentials WHERE client_id=?", (client_id,)
        ).fetchone()
        if remember_password:
            final_enc = new_enc or (old["password_enc"] if old else "")
        else:
            final_enc = ""  # 未勾选保存密码：清除已存密码，但姓名与账号仍保留
        con.execute(
            """INSERT INTO credentials(client_id, submitter, account,
                                       password_enc, updated_at)
               VALUES(?,?,?,?,?)
               ON CONFLICT(client_id) DO UPDATE SET
                   submitter=excluded.submitter,
                   account=excluded.account,
                   password_enc=excluded.password_enc,
                   updated_at=excluded.updated_at""",
            (client_id, submitter, account, final_enc, _now()),
        )


def _clear_credentials(client_id: str) -> None:
    with DB_LOCK, _db() as con:
        con.execute("DELETE FROM credentials WHERE client_id=?", (client_id,))

# ---------------------------------------------------------------------------
# 请求身份与访问密码
# ---------------------------------------------------------------------------

def _valid_client_id(value: str | None) -> bool:
    return bool(value and re.fullmatch(r"[0-9a-f]{32}", value))


def _client_id() -> str:
    value = request.cookies.get(CLIENT_COOKIE, "")
    if _valid_client_id(value):
        return value
    value = uuid.uuid4().hex
    request.environ["pdd_new_client_id"] = value
    return value


@app.before_request
def _require_link_token():
    if not LINK_TOKEN:
        return None
    supplied = request.args.get("token", "")
    if supplied and hmac.compare_digest(supplied, LINK_TOKEN):
        request.environ["pdd_link_token_ok"] = "1"
        return None
    cookie = request.cookies.get(LINK_TOKEN_COOKIE, "")
    if cookie and hmac.compare_digest(cookie, LINK_TOKEN):
        return None
    return Response("Not found", status=404, mimetype="text/plain")


@app.after_request
def _persist_client_cookie(response):
    value = request.environ.get("pdd_new_client_id")
    if request.environ.get("pdd_link_token_ok"):
        response.set_cookie(
            LINK_TOKEN_COOKIE,
            LINK_TOKEN,
            max_age=365 * 24 * 3600,
            httponly=True,
            samesite="Lax",
            secure=request.is_secure,
        )
    value = request.environ.get("pdd_new_client_id")
    if value:
        response.set_cookie(
            CLIENT_COOKIE,
            value,
            max_age=365 * 24 * 3600,
            httponly=True,
            samesite="Lax",
            secure=request.is_secure,
        )
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "SAMEORIGIN"
    return response


def _authenticated() -> bool:
    return not ACCESS_PASSWORD or bool(session.get("web_authorized"))


def _require_auth():
    if not _authenticated():
        return jsonify({"ok": False, "msg": "请先输入访问密码"}), 401
    return None


@app.route("/api/login", methods=["POST"])
def api_login():
    if not ACCESS_PASSWORD:
        session["web_authorized"] = True
        return jsonify({"ok": True})
    data = request.get_json(force=True, silent=True) or {}
    password = str(data.get("password") or "")
    if hmac.compare_digest(password, ACCESS_PASSWORD):
        session["web_authorized"] = True
        return jsonify({"ok": True})
    return jsonify({"ok": False, "msg": "访问密码错误"}), 403


@app.route("/api/profile")
def api_profile():
    denied = _require_auth()
    if denied:
        return denied
    return jsonify({"ok": True, **_profile_for(_client_id()),
                    "auth_required": bool(ACCESS_PASSWORD)})


@app.route("/api/profile", methods=["DELETE"])
def api_profile_delete():
    denied = _require_auth()
    if denied:
        return denied
    _clear_credentials(_client_id())
    return jsonify({"ok": True})


# ---------------------------------------------------------------------------
# 任务队列与后台工作线程
# ---------------------------------------------------------------------------

PROGRESS_STEPS = (
    ("===== 店铺：", "正在处理店铺"),
    ("打开电商融系统", "正在登录电商融"),
    ("商户管理：搜索店铺", "正在获取店铺账号"),
    ("拼多多：探测已有登录态", "正在检查拼多多登录态"),
    ("拼多多：打开账密登录页", "正在登录拼多多"),
    ("输入账号密码", "正在登录拼多多"),
    ("等待短信验证码", "等待短信验证码"),
    ("短信验证码", "等待短信验证码"),
    ("打开 发货管理→订单查询", "正在查询订单"),
    ("查询完成", "订单查询完成"),
    ("批量导出 → 生成报表", "正在生成报表"),
    ("轮询等待报表生成", "正在等待报表生成"),
    ("下载中", "正在下载报表"),
    ("报表已下载", "报表已下载"),
)


def _infer_progress(text: str) -> str:
    current = "正在启动任务"
    for line in text.splitlines()[-240:]:
        for needle, label in PROGRESS_STEPS:
            if needle in line:
                current = label
                break
    return current


def _next_cdp_port(client_id: str) -> int:
    return 9350 + (int(client_id[:8], 16) % 500)


def _job_env(job: sqlite3.Row, account: str, password: str) -> dict:
    client_id = job["client_id"]
    env = dict(os.environ)
    env["DS_ACCOUNT"] = account
    env["DS_PASSWORD"] = password
    env["PYTHONIOENCODING"] = "utf-8"
    env["PDD_DS_PROFILE_DIR"] = str(PROFILES_DIR / client_id / "ds")
    env["PDD_CDP_PROFILE_DIR"] = str(PROFILES_DIR / client_id / "pdd")
    env["PDD_CDP_PORT"] = str(_next_cdp_port(client_id))
    env["PDD_DEBUG_DIR"] = str(JOBS_DIR / job["id"] / "debug")
    env["PDD_SESSION_FILE"] = str(PROFILES_DIR / client_id / "SESSION_SHOP.txt")
    return env


def _job_command(job: sqlite3.Row) -> list[str]:
    shops = json.loads(job["shops"] or "[]")
    job_dir = JOBS_DIR / job["id"]
    out_dir = Path(job["out_dir"])
    return [
        sys.executable,
        str(BASE_DIR / "pdd_export.py"),
        *shops,
        "--cdp",
        "--start", job["start_date"],
        "--end", job["end_date"],
        "--out-dir", str(out_dir),
        "--job-dir", str(job_dir),
        "--profile-dir", str(PROFILES_DIR / job["client_id"] / "ds"),
        "--cdp-profile-dir", str(PROFILES_DIR / job["client_id"] / "pdd"),
        "--cdp-port", str(_next_cdp_port(job["client_id"])),
    ]


def _archive_results(job_id: str, results: list) -> None:
    r"""任务完成后，把成功的报表复制一份到 ORDER_ARCHIVE（默认 E:\order）下的任务子目录。

    仅作对外归档整理；网页下载仍从任务目录取文件，删除归档副本不影响下载。
    任何异常都吞掉，绝不影响任务状态本身。
    """
    files = [it for it in (results or [])
             if it.get("status") == "成功" and it.get("file")]
    if not files:
        return
    try:
        dst_root = ORDER_ARCHIVE / job_id
        dst_root.mkdir(parents=True, exist_ok=True)
        copied = 0
        for it in files:
            rel = it.get("file_rel") or ("downloads/" + it["file"])
            src = JOBS_DIR / job_id / rel
            if src.is_file():
                shutil.copy2(src, dst_root / it["file"])
                copied += 1
        try:
            with open(JOBS_DIR / job_id / "run.log", "a", encoding="utf-8") as fh:
                fh.write(f"[archive] {copied}/{len(files)} 个报表已归档到 {dst_root}\n")
        except Exception:
            pass
    except Exception:
        pass


def _claim_next_job() -> sqlite3.Row | None:
    with DB_LOCK, _db() as con:
        con.execute("BEGIN IMMEDIATE")
        row = con.execute(
            "SELECT * FROM jobs WHERE status='排队中' ORDER BY seq LIMIT 1"
        ).fetchone()
        if not row:
            return None
        con.execute(
            """UPDATE jobs SET status='正在执行', started_at=?,
                      progress='正在启动任务', error=''
               WHERE id=?""",
            (_now(), row["id"]),
        )
        return con.execute("SELECT * FROM jobs WHERE id=?", (row["id"],)).fetchone()


def _set_job_state(job_id: str, *, status: str | None = None,
                   progress: str | None = None, finished_at: str | None = None,
                   results: list | None = None, error: str | None = None) -> None:
    fields, values = [], []
    if status is not None:
        fields.append("status=?")
        values.append(status)
    if progress is not None:
        fields.append("progress=?")
        values.append(progress)
    if finished_at is not None:
        fields.append("finished_at=?")
        values.append(finished_at)
    if results is not None:
        fields.append("results=?")
        values.append(json.dumps(results, ensure_ascii=False))
    if error is not None:
        fields.append("error=?")
        values.append(error)
    if status in ("已完成", "部分失败", "失败", "已取消"):
        fields.append("password_enc=''")
    if not fields:
        return
    values.append(job_id)
    with DB_LOCK, _db() as con:
        con.execute(f"UPDATE jobs SET {', '.join(fields)} WHERE id=?", values)


def _read_tail(path: Path, limit: int = 12000) -> str:
    try:
        with open(path, "rb") as f:
            f.seek(0, io.SEEK_END)
            size = f.tell()
            f.seek(max(0, size - limit))
            raw = f.read()
        for enc in ("utf-8", "gbk"):
            try:
                return raw.decode(enc)
            except UnicodeDecodeError:
                pass
        return raw.decode("utf-8", errors="replace")
    except Exception:
        return ""


def _normalize_results(raw: list, job_dir: Path) -> list[dict]:
    normalized = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        out = {
            "shop": str(item.get("shop") or ""),
            "status": str(item.get("status") or ""),
            "orders": item.get("orders", ""),
            "file": "",
            "file_rel": "",
            "error": str(item.get("error") or ""),
        }
        file_value = str(item.get("file") or "")
        if file_value:
            out["file"] = Path(file_value).name
            try:
                file_path = Path(file_value).resolve()
                if file_path.is_relative_to(job_dir.resolve()) and file_path.is_file():
                    out["file_rel"] = str(file_path.relative_to(job_dir.resolve()))
                else:
                    out["error"] = out["error"] or "报表文件不在任务目录内"
                    out["status"] = "失败"
            except Exception:
                out["error"] = out["error"] or "报表文件路径无效"
                out["status"] = "失败"
        normalized.append(out)
    return normalized


def _parse_summary(text: str) -> list[dict]:
    started = False
    results = []
    for line in text.splitlines():
        if "运行汇总" in line:
            started = True
            continue
        if not started:
            continue
        line = line.strip()
        m = re.match(r"^✅\s*(.+?)：成功，([\d,]+) 单，(.+)$", line)
        if m:
            path = Path(m.group(3).strip())
            results.append({
                "shop": m.group(1).strip(),
                "status": "成功",
                "orders": m.group(2).replace(",", ""),
                "file": str(path),
                "file_rel": "",
                "error": "",
            })
            continue
        m = re.match(r"^❌\s*(.+?)：失败（(.*)）\s*$", line)
        if m:
            results.append({
                "shop": m.group(1).strip(),
                "status": "失败",
                "orders": "",
                "file": "",
                "file_rel": "",
                "error": m.group(2),
            })
    return results

def _final_status(results: list[dict], return_code: int) -> str:
    success = sum(1 for r in results if r.get("status") == "成功")
    failed = sum(1 for r in results if r.get("status") == "失败")
    if success and failed:
        return "部分失败"
    if success:
        return "已完成"
    if return_code != 0 or not results:
        return "失败"
    return "已完成"


def _run_job(job: sqlite3.Row) -> None:
    job_id = job["id"]
    job_dir = JOBS_DIR / job_id
    log_path = Path(job["log_path"])
    out_dir = Path(job["out_dir"])
    job_dir.mkdir(parents=True, exist_ok=True)
    out_dir.mkdir(parents=True, exist_ok=True)
    log_path.parent.mkdir(parents=True, exist_ok=True)

    account = job["account"] or ""
    password = _decrypt_text(job["password_enc"] or "")
    if not account or not password:
        _set_job_state(job_id, status="失败", progress="未找到账号密码",
                       finished_at=_now(), error="任务缺少可用密码")
        return

    env = _job_env(job, account, password)
    cmd = _job_command(job)
    proc = None
    with open(log_path, "wb") as log_file:
        try:
            proc = subprocess.Popen(
                cmd,
                cwd=str(BASE_DIR),
                env=env,
                stdin=subprocess.DEVNULL,
                stdout=log_file,
                stderr=subprocess.STDOUT,
            )
            with RUNNING_LOCK:
                RUNNING["job_id"] = job_id
                RUNNING["proc"] = proc

            last_progress = ""
            while proc.poll() is None:
                text = _read_tail(log_path)
                progress = _infer_progress(text)
                if progress != last_progress:
                    _set_job_state(job_id, progress=progress)
                    last_progress = progress
                time.sleep(1.2)

            return_code = proc.returncode
            current = _get_job(job_id)
            if current and current["status"] == "已取消":
                return
            text = _read_tail(log_path, limit=1024 * 1024)
            result_file = job_dir / "results.json"
            results = []
            if result_file.exists():
                try:
                    results = _normalize_results(
                        json.loads(result_file.read_text(encoding="utf-8")),
                        job_dir,
                    )
                except Exception:
                    results = []
            if not results:
                results = _normalize_results(_parse_summary(text), job_dir)

            status = _final_status(results, return_code)
            error = ""
            if status == "失败":
                error = f"任务退出码 {return_code}"
                for item in reversed(results):
                    if item.get("error"):
                        error = item["error"]
                        break
            _set_job_state(
                job_id,
                status=status,
                progress="已完成" if status == "已完成" else
                         ("排队完成，部分店铺失败" if status == "部分失败" else "执行失败"),
                finished_at=_now(),
                results=results,
                error=error,
            )
            _archive_results(job_id, results)
        except Exception as exc:
            current = _get_job(job_id)
            if not current or current["status"] != "已取消":
                _set_job_state(
                    job_id,
                    status="失败",
                    progress="任务启动或执行失败",
                    finished_at=_now(),
                    error=str(exc)[:500],
                )
        finally:
            try:
                with DB_LOCK, _db() as con:
                    # 任务结束即清掉明文账号与密码（凭据表里按勾选另存）
                    con.execute("UPDATE jobs SET password_enc='', account='' WHERE id=?", (job_id,))
            except Exception:
                pass
            with RUNNING_LOCK:
                if RUNNING.get("job_id") == job_id:
                    RUNNING["job_id"] = ""
                    RUNNING["proc"] = None
            if proc and proc.poll() is None:
                try:
                    proc.kill()
                except Exception:
                    pass


_LAST_SWEEP = 0.0


def _purge_old_jobs() -> None:
    r"""清除结束超过 24 小时的任务：DB 记录 + 内部文件一起删，防止任务队列无限堆积。

    E:\order 里的归档副本不受影响（那是给用户长期翻找的）。"""
    cutoff = (datetime.now() - timedelta(hours=24)).strftime("%Y-%m-%d %H:%M:%S")
    ids = []
    with DB_LOCK, _db() as con:
        rows = con.execute(
            "SELECT id FROM jobs WHERE status IN (?,?,?,?) "
            "AND finished_at != '' AND finished_at < ?",
            (*FINISHED_STATUSES, cutoff),
        ).fetchall()
        ids = [r["id"] for r in rows]
        if ids:
            con.execute(
                f"DELETE FROM jobs WHERE id IN ({','.join('?' * len(ids))})", ids)
    for jid in ids:
        shutil.rmtree(JOBS_DIR / jid, ignore_errors=True)
    if ids:
        print(f"[purge] 已清除 {len(ids)} 个过期任务", flush=True)


def _worker_loop() -> None:
    global _LAST_SWEEP
    while True:
        job = _claim_next_job()
        if not job:
            WORKER_WAKE.wait(2.0)
            WORKER_WAKE.clear()
            if time.time() - _LAST_SWEEP > 600:  # 每 10 分钟清扫一次过期任务
                _LAST_SWEEP = time.time()
                try:
                    _purge_old_jobs()
                except Exception as exc:
                    print(f"[purge] 清扫失败: {exc}", flush=True)
            continue
        WORKER_WAKE.clear()
        _run_job(job)


def _start_worker() -> None:
    global WORKER_STARTED
    with WORKER_LOCK:
        if WORKER_STARTED:
            return
        WORKER_STARTED = True
        threading.Thread(target=_worker_loop, name="pdd-job-worker",
                         daemon=True).start()


# ---------------------------------------------------------------------------
# 任务 API
# ---------------------------------------------------------------------------

@app.route("/api/run", methods=["POST"])
def api_run():
    denied = _require_auth()
    if denied:
        return denied
    data = request.get_json(force=True, silent=True) or {}
    shops = [str(s).strip() for s in (data.get("shops") or []) if str(s).strip()]
    shops = list(dict.fromkeys(shops))
    submitter = str(data.get("submitter") or "").strip()
    account = str(data.get("ds_account") or "").strip()
    password = str(data.get("ds_password") or "")
    save_password = bool(data.get("save_password"))
    s = str(data.get("start") or "").strip()
    e = str(data.get("end") or "").strip()

    if not submitter:
        return jsonify({"ok": False, "msg": "请填写使用人姓名"}), 400
    if len(submitter) > 30:
        return jsonify({"ok": False, "msg": "使用人姓名不能超过 30 个字符"}), 400
    if not shops:
        return jsonify({"ok": False, "msg": "请至少填写一个店铺名"}), 400
    if len(shops) > 30:
        return jsonify({"ok": False, "msg": "单个任务最多填写 30 个店铺"}), 400
    try:
        ds = datetime.strptime(s, "%Y-%m-%d")
        de = datetime.strptime(e, "%Y-%m-%d")
    except ValueError:
        return jsonify({"ok": False, "msg": "日期格式应为 YYYY-MM-DD"}), 400
    if de < ds:
        return jsonify({"ok": False, "msg": "结束日期早于开始日期"}), 400
    if (de - ds).days > MAX_BACKOFF_DAYS:
        return jsonify({"ok": False, "msg": "时间窗超过 3 个月，拼多多不允许导出"}), 400

    client_id = _client_id()
    saved_name, saved_account, saved_password = _saved_credentials(client_id)
    submitter = submitter or saved_name
    if not account:
        account = saved_account
    if not password:
        password = saved_password
    if not account:
        return jsonify({"ok": False, "msg": "请填写电商融账号"}), 400
    if not password:
        return jsonify({"ok": False, "msg": "请填写电商融密码"}), 400

    if save_password:
        _upsert_credentials(client_id, submitter, account, password, True)
    else:
        # 未勾选保存密码：只保存使用人姓名与账号（下次免填），清除已存密码
        _upsert_credentials(client_id, submitter, account, "", False)

    job_id = datetime.now().strftime("%Y%m%d_%H%M%S_") + uuid.uuid4().hex[:8]
    job_dir = JOBS_DIR / job_id
    out_dir = job_dir / "downloads"
    job_dir.mkdir(parents=True, exist_ok=True)
    out_dir.mkdir(parents=True, exist_ok=True)
    log_path = job_dir / "run.log"
    now = _now()
    with DB_LOCK, _db() as con:
        con.execute(
            """INSERT INTO jobs(
                   id, client_id, submitter, created_at, status, progress,
                   shops, account, password_enc, start_date, end_date,
                   log_path, out_dir
               ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (job_id, client_id, submitter, now, "排队中", "等待执行",
             json.dumps(shops, ensure_ascii=False), account,
             _encrypt_text(password), s, e, str(log_path), str(out_dir)),
        )
    WORKER_WAKE.set()
    return jsonify({"ok": True, "job_id": job_id, "msg": "已加入队列"})


@app.route("/api/status")
def api_status():
    denied = _require_auth()
    if denied:
        return denied
    client_id = _client_id()
    with RUNNING_LOCK:
        running_id = RUNNING.get("job_id") or ""
    return jsonify({
        "ok": True,
        "jobs": _jobs_for_client(client_id),
        "running_job_id": running_id,
        "auth_required": bool(ACCESS_PASSWORD),
        "server_time": _now(),
    })


# 页面日志降噪：过滤 Traceback 内部帧 / Playwright 调用日志 / node 警告等噪声，
# 只留时间戳行、结论行（✅/❌/⚠️/汇总）。判断结果看 ❌ 行即可，完整细节仍在 run.log 文件里。
_LOG_NOISE = re.compile(
    r"Traceback \(most recent call last\)"
    r"|^\s*File \""
    r"|^\s*\^+"
    r"|^\s*~+"
    r"|^\s*\.{3}<\d+ lines>"
    r"|^\s*\)"
    r"|rewrite_error|_impl_obj|from_channel|task\.result\(\)|self\._sync"
    r"|process_shop\(pdd_ctx|ds_open_shop_modal\(ds_page|box\.click\(\)"
    r"|return from_channel|return task\.result|await self\._channel"
    r"|DeprecationWarning|trace-deprecation|\(node:\d+\)"
    r"|^\s+raise |^\s+return mapping\.|^\s+return self\._sync"
)


def _clean_log_text(text: str, keep: int = 120) -> str:
    kept = [ln for ln in text.splitlines() if not _LOG_NOISE.search(ln)]
    # 压掉连续空行
    out, blank = [], 0
    for ln in kept:
        blank = blank + 1 if not ln.strip() else 0
        if blank <= 1:
            out.append(ln)
    return "\n".join(out[-keep:])


@app.route("/api/jobs/<job_id>/log")
def api_job_log(job_id: str):
    denied = _require_auth()
    if denied:
        return denied
    row = _get_job(job_id)
    if not row:
        return jsonify({"ok": False, "msg": "任务不存在"}), 404
    if row["client_id"] != _client_id():
        return jsonify({"ok": False, "msg": "无权查看该任务日志"}), 403
    text = _read_tail(Path(row["log_path"]), limit=60000)
    return jsonify({"ok": True, "log": _clean_log_text(text)})


@app.route("/api/jobs/<job_id>/cancel", methods=["POST"])
def api_job_cancel(job_id: str):
    denied = _require_auth()
    if denied:
        return denied
    row = _get_job(job_id)
    if not row:
        return jsonify({"ok": False, "msg": "任务不存在"}), 404
    if row["client_id"] != _client_id():
        return jsonify({"ok": False, "msg": "无权取消该任务"}), 403
    if row["status"] == "排队中":
        _set_job_state(job_id, status="已取消", progress="已取消",
                       finished_at=_now())
        return jsonify({"ok": True, "msg": "排队任务已取消"})
    if row["status"] == "正在执行":
        with RUNNING_LOCK:
            proc = RUNNING.get("proc") if RUNNING.get("job_id") == job_id else None
        if proc and proc.poll() is None:
            try:
                subprocess.run(
                    ["taskkill", "/T", "/F", "/PID", str(proc.pid)],
                    capture_output=True,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
            except Exception:
                try:
                    proc.kill()
                except Exception:
                    pass
        _set_job_state(job_id, status="已取消", progress="已取消",
                       finished_at=_now())
        return jsonify({"ok": True, "msg": "正在执行的任务已取消"})
    return jsonify({"ok": False, "msg": "该任务已结束，不能取消"}), 400


def _resolve_result_file(row: sqlite3.Row, index: int) -> Path | None:
    try:
        results = json.loads(row["results"] or "[]")
        item = results[index]
    except Exception:
        return None
    if item.get("status") != "成功" or not item.get("file_rel"):
        return None
    job_dir = (JOBS_DIR / row["id"]).resolve()
    file_path = (job_dir / item["file_rel"]).resolve()
    if not file_path.is_relative_to(job_dir) or not file_path.is_file():
        return None
    return file_path


@app.route("/api/jobs/<job_id>/files/<int:index>")
def api_job_file(job_id: str, index: int):
    denied = _require_auth()
    if denied:
        return denied
    row = _get_job(job_id)
    if not row:
        return jsonify({"ok": False, "msg": "任务不存在"}), 404
    if row["client_id"] != _client_id():
        return jsonify({"ok": False, "msg": "无权下载该报表"}), 403
    file_path = _resolve_result_file(row, index)
    if not file_path:
        return jsonify({"ok": False, "msg": "报表文件不存在或已被清理"}), 404
    return send_file(file_path, as_attachment=True, download_name=file_path.name)


@app.route("/api/jobs/<job_id>/download-all")
def api_job_download_all(job_id: str):
    denied = _require_auth()
    if denied:
        return denied
    row = _get_job(job_id)
    if not row:
        return jsonify({"ok": False, "msg": "任务不存在"}), 404
    if row["client_id"] != _client_id():
        return jsonify({"ok": False, "msg": "无权下载该任务报表"}), 403
    try:
        results = json.loads(row["results"] or "[]")
    except Exception:
        results = []
    files = []
    for idx, _ in enumerate(results):
        p = _resolve_result_file(row, idx)
        if p:
            files.append(p)
    if not files:
        return jsonify({"ok": False, "msg": "没有可下载的报表"}), 404
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in files:
            zf.write(path, arcname=path.name)
    buf.seek(0)
    return send_file(buf, mimetype="application/zip", as_attachment=True,
                     download_name=f"reports_{job_id}.zip")


# ---------------------------------------------------------------------------
# 页面
# ---------------------------------------------------------------------------

PAGE = r"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>拼多多订单导出 · 电商融</title>
<style>
  :root {
    --primary: #4f46e5; --primary-2: #6366f1;
    --bg: #eef1f7; --card: #ffffff; --line: #e2e8f0;
    --text: #0f172a; --text-2: #64748b; --text-3: #94a3b8;
    --danger: #dc2626; --ok: #059669; --warn: #d97706;
    --radius: 14px; --shadow: 0 1px 2px rgba(15,23,42,.04), 0 8px 24px rgba(15,23,42,.06);
  }
  * { box-sizing: border-box; }
  html, body { margin: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "PingFang SC", "Microsoft YaHei",
                 "Segoe UI", Roboto, sans-serif;
    background: radial-gradient(1200px 500px at 20% -10%, #e0e7ff 0%, transparent 60%),
                radial-gradient(900px 400px at 110% 0%, #ede9fe 0%, transparent 55%),
                var(--bg);
    color: var(--text); min-height: 100vh; font-size: 14px;
  }
  header { display:flex; align-items:center; gap:12px;
           padding:14px 28px; background:rgba(255,255,255,.85); backdrop-filter:blur(8px);
           border-bottom:1px solid var(--line); position:sticky; top:0; z-index:10; }
  .logo { width:36px; height:36px; border-radius:10px; flex:0 0 36px;
          background:linear-gradient(135deg,#e11d48,#f97316);
          color:#fff; font-weight:700; font-size:19px; display:flex;
          align-items:center; justify-content:center; box-shadow:0 4px 10px rgba(225,29,72,.28); }
  .brand h1 { font-size:16px; margin:0; letter-spacing:.5px; }
  .brand p { margin:0; font-size:12px; color:var(--text-2); }
  .live { margin-left:auto; display:flex; align-items:center; gap:7px;
          font-size:12px; color:var(--text-2); background:#fff; border:1px solid var(--line);
          padding:5px 12px; border-radius:999px; }
  .live .dot { width:8px; height:8px; border-radius:50%; background:var(--ok);
               box-shadow:0 0 0 0 rgba(5,150,105,.5); animation:pulse 1.8s infinite; }
  @keyframes pulse { 70% { box-shadow:0 0 0 7px rgba(5,150,105,0); } }
  .wrap { max-width:1180px; margin:26px auto 60px; padding:0 24px;
          display:grid; grid-template-columns:430px 1fr; gap:22px; align-items:start; }
  .col-form { position:sticky; top:80px; }
  .card { background:var(--card); border:1px solid var(--line); border-radius:var(--radius);
          padding:22px; box-shadow:var(--shadow); margin-bottom:20px; }
  .card h2 { font-size:15px; margin:0 0 4px; display:flex; align-items:center; gap:8px; }
  .card h2::before { content:""; width:4px; height:15px; border-radius:2px;
                     background:linear-gradient(180deg,#6366f1,#4f46e5); }
  .card .sub { font-size:12px; color:var(--text-3); margin:0 0 14px; }
  label { display:block; font-size:12.5px; color:var(--text-2); margin:14px 0 6px;
          font-weight:500; }
  label .n { color:var(--primary); font-weight:600; }
  textarea, input[type=text], input[type=password], input[type=date] {
      width:100%; padding:10px 12px; border:1.5px solid var(--line); border-radius:9px;
      font-size:14px; font-family:inherit; background:#fbfcfe; color:var(--text);
      transition:border-color .15s, box-shadow .15s, background .15s; }
  textarea { resize:vertical; min-height:96px; line-height:1.7; }
  textarea:focus, input:focus { outline:none; border-color:#6366f1;
      background:#fff; box-shadow:0 0 0 3px rgba(99,102,241,.14); }
  .row2 { display:grid; grid-template-columns:1fr 1fr; gap:12px; }
  .chips { display:flex; gap:8px; margin:8px 0 2px; flex-wrap:wrap; }
  .chip { font-size:12px; color:var(--text-2); background:#f1f5f9; border:1px solid var(--line);
          border-radius:999px; padding:4px 12px; cursor:pointer; user-select:none;
          transition:all .15s; }
  .chip:hover { color:var(--primary); border-color:#6366f1; background:#eef2ff; }
  button.go { width:100%; border:0; border-radius:10px; margin-top:18px;
      padding:13px; font-size:15px; font-weight:600; cursor:pointer; color:#fff;
      background:linear-gradient(135deg,#4f46e5,#6366f1);
      box-shadow:0 6px 16px rgba(79,70,229,.3);
      transition:transform .12s, box-shadow .15s, filter .15s; }
  button.go:hover { filter:brightness(1.06); box-shadow:0 8px 20px rgba(79,70,229,.38); }
  button.go:active { transform:translateY(1px); }
  button.go:disabled { background:#a5b4c8; box-shadow:none; cursor:not-allowed; }
  button.ghost { border:1px solid var(--line); background:#fff; color:var(--text-2);
                 border-radius:8px; padding:7px 14px; font-size:12.5px; cursor:pointer;
                 transition:all .15s; }
  button.ghost:hover { color:var(--danger); border-color:#fecaca; background:#fef2f2; }
  .tip { font-size:12px; color:var(--text-2); margin-top:14px; line-height:1.9;
         background:#f8fafc; border:1px dashed var(--line); border-radius:10px; padding:10px 14px; }
  .save-row { margin-top:14px; display:flex; gap:10px; align-items:center; flex-wrap:wrap;
              font-size:12.5px; color:var(--text-2); }
  .save-row label { margin:0; display:inline-flex; gap:6px; align-items:center; cursor:pointer; }
  .save-row input[type=checkbox] { accent-color:var(--primary); width:15px; height:15px; }
  .status-line { font-size:12px; color:var(--text-3); }
  .table-wrap { max-height:400px; overflow:auto; border:1px solid var(--line);
                border-radius:10px; margin-top:12px; }
  table { width:100%; border-collapse:collapse; font-size:13px; min-width:640px; }
  th, td { padding:10px 12px; text-align:left; vertical-align:top;
           border-bottom:1px solid #eef2f6; }
  th { font-size:12px; color:var(--text-2); font-weight:600; background:#f8fafc;
       position:sticky; top:0; z-index:2; white-space:nowrap; }
  tbody tr { transition:background .12s; cursor:pointer; }
  tbody tr:hover { background:#f8faff; }
  tbody tr.running { background:linear-gradient(90deg,#fffbeb,#fff); }
  tbody tr.running td:first-child { box-shadow:inset 3px 0 0 var(--warn); }
  tbody tr.selected { background:#eef2ff; }
  tbody tr.selected td:first-child { box-shadow:inset 3px 0 0 var(--primary); }
  td .shops { font-weight:500; }
  td .date { font-size:12px; color:var(--text-3); margin-top:3px; }
  td .prog { font-size:12.5px; }
  td .time { font-size:11.5px; color:var(--text-3); margin-top:3px; }
  .tag { display:inline-flex; align-items:center; gap:5px; border-radius:999px;
         padding:3px 10px; font-size:12px; white-space:nowrap; font-weight:500; }
  .tag i { width:6px; height:6px; border-radius:50%; background:currentColor; flex:0 0 6px; }
  .tag.queued { background:#f1f5f9; color:#64748b; }
  .tag.running { background:#fef3c7; color:#b45309; }
  .tag.running i { animation:blink 1s infinite; }
  @keyframes blink { 50% { opacity:.25; } }
  .tag.done { background:#ecfdf5; color:var(--ok); }
  .tag.fail { background:#fef2f2; color:var(--danger); }
  .tag.partial { background:#fff7ed; color:var(--warn); }
  .tag.cancel { background:#f1f5f9; color:#94a3b8; }
  .muted { color:var(--text-3); }
  .link { color:var(--primary); text-decoration:none; margin-right:8px; font-size:12.5px; }
  .link:hover { text-decoration:underline; }
  a.btn-dl { display:inline-flex; align-items:center; gap:5px; background:#ecfdf5;
             color:var(--ok); border-radius:8px; padding:5px 11px; font-size:12.5px;
             text-decoration:none; font-weight:600; border:1px solid #a7f3d0;
             transition:all .15s; }
  a.btn-dl:hover { background:#d1fae5; }
  button.mini { border:1px solid var(--line); background:#fff; color:var(--text-2);
                border-radius:7px; padding:4px 10px; font-size:12px; cursor:pointer; }
  button.mini.danger { color:var(--danger); border-color:#fecaca; }
  button.mini.danger:hover { background:#fef2f2; }
  .empty { text-align:center; padding:34px 0; color:var(--text-3); font-size:13px; }
  .empty .big { font-size:30px; display:block; margin-bottom:8px; opacity:.45; }
  .terminal { background:#0f172a; color:#cbd5e1; padding:16px 18px; border-radius:10px;
      font-family:Consolas, "Courier New", monospace; font-size:12px; line-height:1.75;
      max-height:330px; overflow:auto; white-space:pre-wrap; word-break:break-all; }
  .terminal::-webkit-scrollbar { width:8px; height:8px; }
  .terminal::-webkit-scrollbar-thumb { background:#334155; border-radius:4px; }
  .login-overlay { position:fixed; inset:0; background:rgba(15,23,42,.55);
                   backdrop-filter:blur(4px); display:flex; align-items:center;
                   justify-content:center; z-index:20; }
  .login-box { width:min(400px,92vw); background:#fff; border-radius:16px; padding:30px;
               box-shadow:0 25px 60px rgba(15,23,42,.35); border-top:4px solid var(--primary); }
  .login-box h2 { margin:0 0 4px; font-size:19px; }
  .login-box .sub { font-size:12.5px; color:var(--text-2); margin-bottom:16px; }
  .login-box input { margin-top:4px; }
  .hidden { display:none !important; }
  @media (max-width: 1020px) {
    .wrap { grid-template-columns:1fr; }
    .col-form { position:static; }
  }
  @media (max-width: 640px) {
    body { font-size:13px; } .wrap { padding:0 14px; margin-top:16px; }
    header { padding:12px 16px; } .row2 { grid-template-columns:1fr; }
    .card { padding:16px; }
  }
</style>
</head>
<body>
<div id="loginOverlay" class="login-overlay hidden">
  <div class="login-box">
    <h2>拼多多订单导出</h2>
    <div class="sub">输入访问密码以继续</div>
    <label>访问密码</label>
    <input type="password" id="accessPassword" autocomplete="current-password"
           onkeydown="if(event.key==='Enter')login()">
    <div style="margin-top:16px"><button class="go" style="margin-top:0" onclick="login()">进 入</button></div>
    <div id="loginError" class="status-line" style="color:var(--danger);margin-top:10px"></div>
  </div>
</div>

<header>
  <div class="logo">拼</div>
  <div class="brand">
    <h1>拼多多订单导出</h1>
    <p>电商融 · 多人协作队列</p>
  </div>
  <div class="live"><span class="dot"></span>服务运行中 · 每 2 秒自动刷新</div>
</header>

<div class="wrap">
  <div class="col-form">
    <div class="card">
      <h2>新建导出任务</h2>
      <p class="sub">填写后提交，任务进入服务器队列串行执行</p>

      <label>使用人姓名</label>
      <input type="text" id="submitter" maxlength="30" placeholder="例如：张三">

      <label>店铺名称（每行一个，可填多家）<span class="n" id="shopCount"></span></label>
      <textarea id="shops" placeholder="例如：&#10;示例旗舰店&#10;示例水果专卖店"
                oninput="countShops()"></textarea>

      <div class="row2">
        <div><label>电商融账号</label>
          <input type="text" id="ds_account" autocomplete="username"
                 placeholder="手机号 / 用户名"></div>
        <div><label>电商融密码</label>
          <input type="password" id="ds_password" autocomplete="current-password"
                 placeholder="未保存"></div>
      </div>
      <div class="save-row">
        <label><input type="checkbox" id="save_password" checked>同时记住密码（使用人与账号始终自动记住）</label>
        <button class="ghost" type="button" onclick="clearSavedPassword()">清除已保存</button>
        <span id="savedStatus" class="status-line"></span>
      </div>

      <label>导出日期范围（按下单时间）</label>
      <div class="row2">
        <div><input type="date" id="start"></div>
        <div><input type="date" id="end"></div>
      </div>
      <div class="chips">
        <span class="chip" onclick="quickRange(1)">昨天</span>
        <span class="chip" onclick="quickRange(3)">近 3 天</span>
        <span class="chip" onclick="quickRange(7)">近 7 天</span>
        <span class="chip" onclick="quickRange(30)">近 30 天</span>
      </div>

      <button id="go" class="go" onclick="startJob()">开始导出 →</button>

      <div class="tip">
        · 同一时间只执行一个任务，提交后自动排队，可在下方看到实时进度。<br>
        · 报表完成后点「下载报表」，文件经浏览器下载到你电脑；服务器同时自动留档。<br>
        · 拼多多两次生成报表间隔约 5 分钟，多店铺串行时程序会自动等待，请耐心刷新。
      </div>
    </div>
  </div>

  <div class="col-main">
    <div class="card">
      <div class="head-row">
        <h2>任务队列</h2>
        <span class="status-line">已结束任务保留 1 小时</span>
      </div>
      <div class="table-wrap">
        <table>
          <thead><tr>
            <th style="width:64px">位置</th>
            <th style="width:96px">使用人</th>
            <th>店铺 / 日期</th>
            <th style="width:104px">状态</th>
            <th>当前步骤</th>
            <th style="width:168px">操作</th>
          </tr></thead>
          <tbody id="jobRows">
            <tr><td colspan="6"><div class="empty"><span class="big">📭</span>暂无任务，在左侧提交第一个导出任务</div></td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <div class="card">
      <div class="head-row">
        <h2>运行日志</h2>
        <span class="status-line" id="logTitle">未选择任务</span>
      </div>
      <pre id="log" class="terminal">等待任务…</pre>
    </div>
  </div>
</div>

<script>
const $ = id => document.getElementById(id);
let selectedJobId = "";
let loaded = false;

function escapeHtml(value) {
  return String(value == null ? "" : value)
    .replaceAll("&", "&amp;").replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;").replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function statusClass(status) {
  if (status === "正在执行") return "running";
  if (status === "排队中") return "queued";
  if (status === "已完成") return "done";
  if (status === "部分失败") return "partial";
  if (status === "已取消") return "cancel";
  return "fail";
}

function todayDefaults() {
  const today = new Date();
  const end = new Date(today.getTime() - 864e5);
  const start = new Date(today.getTime() - 2 * 864e5);
  $("start").value = start.toISOString().slice(0, 10);
  $("end").value = end.toISOString().slice(0, 10);
}

function quickRange(days) {
  const end = new Date(Date.now() - 864e5);
  const start = new Date(Date.now() - days * 864e5);
  $("start").value = start.toISOString().slice(0, 10);
  $("end").value = end.toISOString().slice(0, 10);
}

function countShops() {
  const n = $("shops").value.split("\n").map(x => x.trim()).filter(Boolean).length;
  $("shopCount").textContent = n ? ("（已填 " + n + " 家）") : "";
}

async function api(path, opts) {
  const response = await fetch(path, opts || {});
  let data = {};
  try { data = await response.json(); } catch (_) {}
  if (response.status === 401 && data.msg === "请先输入访问密码") {
    showLogin(true);
  }
  return {response, data};
}

function showLogin(show) {
  $("loginOverlay").classList.toggle("hidden", !show);
}

async function login() {
  const password = $("accessPassword").value;
  const {response, data} = await api("/api/login", {
    method: "POST",
    headers: {"Content-Type": "application/json"},
    body: JSON.stringify({password})
  });
  if (!response.ok || !data.ok) {
    $("loginError").textContent = data.msg || "登录失败";
    return;
  }
  $("loginError").textContent = "";
  showLogin(false);
  await loadProfile();
  await poll();
}

async function loadProfile() {
  const {response, data} = await api("/api/profile");
  if (!response.ok) return;
  if (data.submitter) $("submitter").value = data.submitter;
  if (data.account) $("ds_account").value = data.account;
  if (data.has_password) {
    $("save_password").checked = true;
    $("ds_password").placeholder = "已保存，留空继续使用";
    $("savedStatus").textContent = "已保存密码";
  } else {
    $("ds_password").placeholder = "未保存";
    $("savedStatus").textContent = "";
  }
  loaded = true;
}

async function clearSavedPassword() {
  if (!confirm("确定清除当前浏览器保存的电商融密码吗？")) return;
  const {response, data} = await api("/api/profile", {method: "DELETE"});
  if (!response.ok || !data.ok) {
    alert(data.msg || "清除失败");
    return;
  }
  $("ds_password").value = "";
  $("ds_password").placeholder = "未保存";
  $("save_password").checked = false;
  $("savedStatus").textContent = "";
}

async function startJob() {
  const shops = $("shops").value.split("\n").map(x => x.trim()).filter(Boolean);
  const body = {
    submitter: $("submitter").value.trim(),
    shops,
    ds_account: $("ds_account").value.trim(),
    ds_password: $("ds_password").value,
    save_password: $("save_password").checked,
    start: $("start").value,
    end: $("end").value
  };
  $("go").disabled = true;
  try {
    const {response, data} = await api("/api/run", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify(body)
    });
    if (!response.ok || !data.ok) {
      alert(data.msg || "提交失败");
      return;
    }
    selectedJobId = data.job_id;
    if ($("save_password").checked) {
      $("ds_password").value = "";
      $("ds_password").placeholder = "已保存，留空继续使用";
      $("savedStatus").textContent = "已保存密码";
    }
    await poll();
  } finally {
    $("go").disabled = false;
  }
}

async function cancelJob(jobId) {
  if (!confirm("确定取消这个任务吗？")) return;
  const {data} = await api(`/api/jobs/${encodeURIComponent(jobId)}/cancel`, {method: "POST"});
  if (!data.ok) alert(data.msg || "取消失败");
  await poll();
}

function renderJobs(jobs) {
  const tbody = $("jobRows");
  if (!jobs.length) {
    tbody.innerHTML = '<tr><td colspan="6"><div class="empty"><span class="big">📭</span>暂无任务，在左侧提交第一个导出任务</div></td></tr>';
    return;
  }
  const rows = [];
  for (const job of jobs) {
    const cls = job.status === "正在执行" ? "running" : "";
    const selected = job.id === selectedJobId ? "selected" : "";
    const queue = job.status === "排队中" && job.queue_position
      ? "排队中 · 第 " + job.queue_position + " 位" : job.status;
    const shops = (job.shops || []).map(escapeHtml).join("、");
    const dateRange = escapeHtml(job.start_date) + " ~ " + escapeHtml(job.end_date);
    const operations = [];
    if (job.is_owner && (job.status === "排队中" || job.status === "正在执行")) {
      operations.push('<button class="mini ' + (job.status === "正在执行" ? "danger" : "") +
        '" onclick="cancelJob(\'' + escapeHtml(job.id) + '\')">取消</button>');
    }
    if (job.is_owner && job.status === "已完成") {
      const urls = (job.results || []).filter(r => r.download_url).map(r => r.download_url);
      if (urls.length === 1) {
        operations.push('<a class="btn-dl" href="' + escapeHtml(urls[0]) + '">⬇ 下载报表</a>');
      } else if (urls.length > 1) {
        operations.push('<a class="btn-dl" href="/api/jobs/' + encodeURIComponent(job.id) +
          '/download-all">⬇ 打包下载 ' + urls.length + ' 份</a>');
      }
    }
    operations.push('<a class="link" href="#" onclick="selectJob(\'' + escapeHtml(job.id) +
      '\');return false">日志</a>');
    rows.push('<tr class="' + cls + ' ' + selected + '" onclick="selectJob(\'' +
      escapeHtml(job.id) + '\')">' +
      '<td>' + (job.status === "排队中" ? (job.queue_position || "-") : "-") + '</td>' +
      '<td>' + escapeHtml(job.submitter) + '</td>' +
      '<td><div class="shops">' + shops + '</div><div class="date">📅 ' + dateRange + '</div></td>' +
      '<td><span class="tag ' + statusClass(job.status) + '"><i></i>' + escapeHtml(queue) + '</span></td>' +
      '<td><div class="prog">' + escapeHtml(job.progress || "") + '</div><div class="time">' +
      escapeHtml(job.created_at) + '</div></td>' +
      '<td>' + operations.join(" ") + '</td></tr>');
  }
  tbody.innerHTML = rows.join("");
}

function selectJob(jobId) {
  selectedJobId = jobId;
  loadLog();
}

async function loadLog() {
  if (!selectedJobId) return;
  const {response, data} = await api(`/api/jobs/${encodeURIComponent(selectedJobId)}/log`);
  if (!response.ok) {
    pre.textContent = "这条任务属于其他同事，日志与文件仅提交人可见";
    return;
  }
  const pre = $("log");
  const atBottom = pre.scrollTop + pre.clientHeight >= pre.scrollHeight - 20;
  pre.textContent = data.log || "暂无日志";
  if (atBottom) pre.scrollTop = pre.scrollHeight;
  $("logTitle").textContent = "任务 " + selectedJobId;
}

async function poll() {
  const {response, data} = await api("/api/status");
  if (!response.ok) return;
  if (!loaded) await loadProfile();
  const jobs = data.jobs || [];
  if (!selectedJobId && jobs.length) selectedJobId = jobs[0].id;
  renderJobs(jobs);
  await loadLog();
}

todayDefaults();
setInterval(poll, 2000);
poll();
</script>
</body>
</html>
"""


@app.route("/")
def index():
    return Response(PAGE, mimetype="text/html; charset=utf-8")


def _lan_ip() -> str:
    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def _open_browser() -> None:
    time.sleep(1.2)
    webbrowser.open(f"http://127.0.0.1:{os.environ.get('PDD_WEB_PORT', '8765')}")


def main() -> None:
    _init_db()
    _start_worker()
    host = os.environ.get("PDD_WEB_HOST", "0.0.0.0")
    port = int(os.environ.get("PDD_WEB_PORT", "8765"))
    if os.environ.get("PDD_OPEN_BROWSER", "1") == "1":
        threading.Thread(target=_open_browser, daemon=True).start()
    print("电商融·拼多多订单导出（服务器版）已启动：")
    print(f"  本机访问：      http://127.0.0.1:{port}")
    print(f"  局域网同事访问：http://{_lan_ip()}:{port}")
    if ACCESS_PASSWORD:
        print("  访问密码：      已启用（PDD_WEB_ACCESS_PASSWORD）")
    else:
        print("  访问密码：      未启用，建议在内网服务上设置 PDD_WEB_ACCESS_PASSWORD")
    print("Ctrl+C 退出")
    try:
        from waitress import serve
        serve(app, host=host, port=port, threads=8)
    except ImportError:
        app.run(host=host, port=port, debug=False, threaded=True)


if __name__ == "__main__":
    main()
