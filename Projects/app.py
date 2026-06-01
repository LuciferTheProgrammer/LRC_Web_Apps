# app.py
# -----------------------------------------------------------------------------
# Active Directory User Creation Flask Web App
#
# This file is the main browser-based web application for the Active Directory
# user creation workflow. It displays the upload form, receives the spreadsheet
# and administrator credentials, validates the target directory information,
# starts the backend worker process, streams live job output to the browser, and
# builds the final ZIP download bundle after the job finishes.
#
# The actual Active Directory account creation work is handled by worker.py. This
# Flask app prepares the job, saves the uploaded spreadsheet to a temporary job
# folder, writes job settings to a JSON file, launches worker.py as a subprocess,
# and exposes routes for the live console, result check, and download.
#
# SECURITY / PUBLIC REPO NOTE:
# This copy keeps the same code structure and workflow, but uses dummy placeholder
# values for environment-specific information such as LDAP server names, domain
# prefixes, OU distinguished names, base DNs, and temporary passwords. Replace
# placeholders only in a private/internal deployment environment.
#
# Main routes:
# - GET  /              renders the upload form
# - POST /run           validates input and starts a background job
# - GET  /job/<job_id>  displays the live job console page
# - GET  /stream/<id>   streams worker output using Server-Sent Events
# - GET  /result/<id>   returns final job status and prepares ZIP download
# - GET  /download/<t>  downloads the generated ZIP bundle
# -----------------------------------------------------------------------------

import os
import io
import sys
import json
import zipfile
import datetime
import secrets
import time
import tempfile
import subprocess
import threading
from queue import Queue, Empty

from flask import (
    Flask, request, send_file, render_template, abort,
    redirect, url_for, make_response, Response
)
from ldap3 import Server, Connection, ALL, SUBTREE, BASE
from ldap3.utils.conv import escape_filter_chars

# ---------------------------
# Config
# ---------------------------
ALLOWED_EXTS = {".xlsx", ".xls", ".csv"}
MAX_FILE_MB = 20

LDAP_SERVER = os.environ.get("LDAP_SERVER", "dc01.example.local")
DEFAULT_DOMAIN_PREFIX = os.environ.get("DOMAIN_PREFIX", "example\\")
DEFAULT_TEMP_OU_DN = os.environ.get("DEFAULT_TEMP_OU_DN", "OU=Temp,OU=Users,DC=example,DC=local")
DEFAULT_TEMP_PW = os.environ.get("DEFAULT_TEMP_PW", "ChangeMe123!")
BASE_DN_FALLBACK = os.environ.get("BASE_DN", "DC=example,DC=local")

# NEW: destination OU for post-move (can override via env)
DEST_OU_DN = os.environ.get("DEST_OU_DN", "OU=Active,OU=Users,DC=example,DC=local")

# In-memory caches
DOWNLOADS = {}               # token -> {data,name,ts}
DOWNLOAD_TTL_SEC = 10 * 60

JOBS = {}                    # job_id -> {proc, queue, out_dir, ts, done, returncode, token}
JOB_TTL_SEC = 60 * 30        # keep jobs for 30 min after start

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_FILE_MB * 1024 * 1024
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0

# ---------------------------
# Global no-cache headers
# ---------------------------
@app.after_request
# Adds no-cache headers so browsers do not reuse old job pages or downloads.
def add_no_cache(resp):
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0, proxy-revalidate, private"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp

# ---------------------------
# Housekeeping
# ---------------------------
# Removes expired ZIP downloads from memory.
def cleanup_downloads():
    now = time.time()
    for t in list(DOWNLOADS.keys()):
        if now - DOWNLOADS[t]["ts"] > DOWNLOAD_TTL_SEC:
            DOWNLOADS.pop(t, None)

# Removes old job records and stops any stale worker process.
def cleanup_jobs():
    now = time.time()
    for jid, meta in list(JOBS.items()):
        if now - meta["ts"] > JOB_TTL_SEC:
            try:
                if meta.get("proc") and meta["proc"].poll() is None:
                    meta["proc"].kill()
            except Exception:
                pass
            JOBS.pop(jid, None)

# ---------------------------
# LDAP helpers (no pyad here)
# ---------------------------
# Builds a domain-qualified username when the user enters only a short username.
def build_full_username(username, domain_prefix):
    u = (username or "").strip()
    if "\\" in u or "@" in u:
        return u
    return f"{domain_prefix}{u}" if domain_prefix else u

# Tests whether the supplied admin credentials can bind to LDAP.
def test_bind(ldap_server, full_username, password):
    srv = Server(ldap_server, get_info=ALL)
    return Connection(srv, full_username, password, auto_bind=True)

# Creates and returns an authenticated LDAP connection.
def connect_ldap(ldap_server, full_username, password):
    srv = Server(ldap_server, get_info=ALL)
    return Connection(srv, full_username, password, auto_bind=True)

# Builds a fallback base DN from the LDAP server hostname.
def derive_dn_from_server(ldap_server: str) -> str:
    host = (ldap_server or "").split(":")[0]
    parts = host.split(".")
    if len(parts) >= 2:
        domain = parts[-2:]
        return ",".join(f"DC={p}" for p in domain)
    return BASE_DN_FALLBACK

# Finds the AD base DN using configured fallback values or LDAP root data.
def get_base_dn(conn: Connection, ldap_server: str) -> str:
    if BASE_DN_FALLBACK and BASE_DN_FALLBACK.strip().lower().startswith("dc="):
        return BASE_DN_FALLBACK
    try:
        other = getattr(conn.server.info, "other", None)
        if other:
            dnc = other.get("defaultNamingContext", [])
            if dnc:
                return dnc[0]
    except Exception:
        pass
    try:
        conn.search("", "(objectClass=*)", search_scope=BASE, attributes=["defaultNamingContext"])
        if conn.entries:
            val = conn.entries[0].entry_attributes_as_dict.get("defaultNamingContext")
            if val:
                return val[0]
    except Exception:
        pass
    try:
        conn.search("", "(objectClass=*)", search_scope=BASE, attributes=["namingContexts"])
        if conn.entries:
            ncs = conn.entries[0].entry_attributes_as_dict.get("namingContexts", [])
            for nc in ncs:
                if str(nc).lower().startswith("dc="):
                    return str(nc)
    except Exception:
        pass
    try:
        return derive_dn_from_server(ldap_server)
    except Exception:
        pass
    return "DC=example,DC=local"

# Resolves a typed OU name or full distinguished name into an actual OU DN.
def resolve_ou_dn(user_input, ldap_server, full_username, password):
    user_input = (user_input or "").strip()
    try:
        conn = connect_ldap(ldap_server, full_username, password)
    except Exception:
        return None, []
    try:
        base = get_base_dn(conn, ldap_server)
        if user_input.upper().startswith(("OU=", "CN=")):
            flt = f"(distinguishedName={escape_filter_chars(user_input)})"
            conn.search(search_base=base, search_filter=flt, search_scope=SUBTREE, attributes=["distinguishedName"])
            if conn.entries:
                return conn.entries[0].distinguishedName.value, []
            last_seg = user_input.split(",")[0]
            if last_seg.upper().startswith("OU="):
                user_input = last_seg.split("=", 1)[1]
        flt = f"(&(objectClass=organizationalUnit)(ou={escape_filter_chars(user_input)}))"
        conn.search(base, flt, SUBTREE, attributes=["distinguishedName"])
        if len(conn.entries) == 1:
            return conn.entries[0].distinguishedName.value, []
        if len(conn.entries) > 1:
            return None, [e.distinguishedName.value for e in conn.entries][:10]
        flt = f"(&(objectClass=organizationalUnit)(ou=*{escape_filter_chars(user_input)}*))"
        conn.search(base, flt, SUBTREE, attributes=["distinguishedName"])
        if len(conn.entries) == 1:
            return conn.entries[0].distinguishedName.value, []
        return None, [e.distinguishedName.value for e in conn.entries][:10]
    finally:
        try:
            conn.unbind()
        except Exception:
            pass

# ---------------------------
# Worker management
# ---------------------------
# Starts worker.py as a background process and returns the process plus output queue.
def start_worker(job_json_path, out_dir):
    worker_script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "worker.py")
    q = Queue()
    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"
    env.setdefault("PYTHONUTF8", "1")          # ensure UTF-8 in child
    env.setdefault("PYTHONIOENCODING", "utf-8")

    proc = subprocess.Popen(
        [sys.executable, worker_script, job_json_path],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        bufsize=1,
        universal_newlines=True,
        env=env,
        encoding="utf-8",                      # decode stdout/stderr as UTF-8
        errors="replace"                       # never die on odd bytes
    )

    # Reads worker stdout and stores each line in the queue for live streaming.
    def _reader():
        try:
            for line in proc.stdout:
                q.put(line.rstrip("\r\n"))
        except Exception as e:
            q.put(f"[reader-error] {e}")
        finally:
            proc.wait()
            try:
                err = proc.stderr.read()
                if err:
                    with open(os.path.join(out_dir, "WORKER_STDERR.txt"), "w", encoding="utf-8") as fh:
                        fh.write(err)
            except Exception:
                pass

    t = threading.Thread(target=_reader, daemon=True)
    t.start()
    return proc, q

# ---------------------------
# Routes
# ---------------------------
@app.get("/")
# Renders the main upload page and displays any validation message.
def index():
    cleanup_downloads(); cleanup_jobs()

    code = request.args.get("code", "").strip()
    prefill_user = request.args.get("u", "").strip()

    code_map = {
        "missingcreds": "Enter your AD admin username and password.",
        "badcred": "Invalid AD admin username or password. Please try again.",
        "badou": "Could not resolve the Target OU. Paste a valid DN or try a different OU name.",
        "nofile": "No file uploaded.",
        "badtype": "Unsupported file type. Please upload .xlsx, .xls, or .csv."
    }
    toast = code_map.get(code) if code else None

    return render_template(
        "index.html",
        error=None,
        success=None,
        download_url=None,
        suggestions=[],
        admin_username=prefill_user,
        domain_prefix=DEFAULT_DOMAIN_PREFIX,
        temp_ou_dn=DEFAULT_TEMP_OU_DN,
        ldap_server=LDAP_SERVER,
        toast=toast,
        toast_code=code
    )

@app.post("/run")
# Validates the form, saves the spreadsheet, creates a job file, and starts the worker.
def run():
    cleanup_downloads(); cleanup_jobs()

    admin_username = (request.form.get("admin_username") or "").strip()
    admin_password = (request.form.get("admin_password") or "")
    domain_prefix  = (request.form.get("domain_prefix") or DEFAULT_DOMAIN_PREFIX).strip()
    ou_input       = (request.form.get("temp_ou_dn") or DEFAULT_TEMP_OU_DN).strip()

    if not admin_username or not admin_password:
        return redirect(url_for("index", code="missingcreds"))

    full_username = build_full_username(admin_username, domain_prefix)

    try:
        conn = test_bind(LDAP_SERVER, full_username, admin_password)
        conn.unbind()
    except Exception:
        return redirect(url_for("index", code="badcred", u=admin_username))

    resolved_dn, _ = resolve_ou_dn(ou_input, LDAP_SERVER, full_username, admin_password)
    if not resolved_dn:
        return redirect(url_for("index", code="badou", u=admin_username))

    f = request.files.get("file")
    if not f:
        return redirect(url_for("index", code="nofile", u=admin_username))
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ALLOWED_EXTS:
        return redirect(url_for("index", code="badtype", u=admin_username))

    td = tempfile.mkdtemp(prefix="adjob_")
    in_path = os.path.join(td, "input" + ext)
    f.save(in_path)
    out_dir = os.path.join(td, "out"); os.makedirs(out_dir, exist_ok=True)

    job = {
        "input_path": in_path,
        "target_ou_dn": resolved_dn,
        "ldap_server": LDAP_SERVER,
        "full_username": full_username,
        "password": admin_password,
        "default_temp_pw": DEFAULT_TEMP_PW,
        "out_dir": out_dir,
        "base_dn_fallback": BASE_DN_FALLBACK,
        # NEW: pass destination OU
        "dest_ou_dn": DEST_OU_DN,
    }
    job_path = os.path.join(td, "job.json")
    with open(job_path, "w", encoding="utf-8") as jf:
        json.dump(job, jf)

    job_id = secrets.token_urlsafe(16)
    proc, q = start_worker(job_path, out_dir)
    JOBS[job_id] = {
        "proc": proc,
        "queue": q,
        "out_dir": out_dir,
        "ts": time.time(),
        "done": False,
        "returncode": None,
        "token": None
    }

    # Waits for the worker process to finish and stores the return code.
    def _waiter(jid):
        p = JOBS[jid]["proc"]
        rc = p.wait()
        JOBS[jid]["returncode"] = rc
        JOBS[jid]["done"] = True
    threading.Thread(target=_waiter, args=(job_id,), daemon=True).start()

    return redirect(url_for("job_page", job_id=job_id))

@app.get("/job/<job_id>")
# Renders the live job console page for a specific job.
def job_page(job_id):
    cleanup_downloads(); cleanup_jobs()
    if job_id not in JOBS:
        abort(404)
    return render_template("job.html", job_id=job_id)

@app.get("/stream/<job_id>")
# Streams job output to the browser using Server-Sent Events.
def stream(job_id):
    cleanup_jobs()
    meta = JOBS.get(job_id)
    if not meta:
        abort(404)
    q = meta["queue"]

    # Yields queued log lines and keep-alive messages for the SSE stream.
    def gen():
        while True:
            try:
                line = q.get(timeout=1.0)
                yield f"data: {line}\n\n"
            except Empty:
                yield ": keep-alive\n\n"
            if meta["done"] and q.empty():
                rc = meta["returncode"]
                yield f"event: done\ndata: {rc}\n\n"
                break

    # Explicitly declare UTF-8 for SSE
    resp = Response(gen(), content_type="text/event-stream; charset=utf-8")
    resp.headers["Cache-Control"] = "no-store"
    resp.headers["X-Accel-Buffering"] = "no"
    return resp

@app.get("/result/<job_id>")
# Returns the final job result and prepares the ZIP download bundle.
def result(job_id):
    cleanup_downloads(); cleanup_jobs()
    meta = JOBS.get(job_id)
    if not meta:
        return {"done": True, "success": False, "error": "unknown job"}, 404

    done = bool(meta["done"])
    if not done:
        return {"done": False}

    if not meta.get("token"):
        out_dir = meta["out_dir"]
        result_path = os.path.join(out_dir, "result.json")
        count = 0
        ok = False
        if os.path.exists(result_path):
            try:
                with open(result_path, "r", encoding="utf-8") as rf:
                    rj = json.load(rf)
                    ok = bool(rj.get("ok"))
                    count = int(rj.get("count", 0))
            except Exception:
                ok = meta["returncode"] == 0
        else:
            ok = meta["returncode"] == 0

        zipbuf = io.BytesIO()
        with zipfile.ZipFile(zipbuf, "w", zipfile.ZIP_DEFLATED) as z:
            # include all worker outputs
            for name in (
                "summary.xlsx",
                "run_log.txt",
                "ERROR.txt",
                "WORKER_STDERR.txt",
                "post_process.ps1",
                "created_users.json",
                "AssignOffice365E1.ps1",   # filename kept
            ):
                p = os.path.join(out_dir, name)
                if os.path.exists(p):
                    with open(p, "rb") as fh:
                        z.writestr(name, fh.read())

            # helpful readme
            readme = (
                "Bundle contents:\n"
                "- summary.xlsx : emails of processed users (Sheet1)\n"
                "- post_process.ps1 : AD post-step script executed during the run\n"
                "- AssignOffice365E1.ps1 : Right-click > Run with PowerShell to assign Office 365 E1 licenses using summary.xlsx\n\n"
                "Note: Microsoft 365 directory sync can take up to ~30 minutes to reflect new proxy addresses & moves.\n"
            )
            z.writestr("README.txt", readme)

        zipbuf.seek(0)

        token = secrets.token_urlsafe(16)
        DOWNLOADS[token] = {
            "data": zipbuf.getvalue(),
            "name": f"ad_user_creation_{datetime.datetime.now().strftime('%Y%m%d-%H%M%S')}.zip",
            "ts": time.time()
        }
        meta["token"] = token
        meta["count"] = count
        meta["success"] = ok

    return {
        "done": True,
        "success": bool(meta.get("success")),
        "count": int(meta.get("count", 0)),
        "download_url": f"/download/{meta['token']}"
    }

@app.get("/download/<token>")
# Sends the generated ZIP file to the browser.
def download(token):
    cleanup_downloads()
    meta = DOWNLOADS.get(token)
    if not meta:
        abort(404)

    data = meta["data"]
    name = meta["name"]

    resp = make_response(send_file(
        io.BytesIO(data),
        mimetype="application/zip",
        as_attachment=True,
        download_name=name,
        max_age=0,
        conditional=False
    ))
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0, proxy-revalidate, private"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5050, threaded=True, use_reloader=False)
