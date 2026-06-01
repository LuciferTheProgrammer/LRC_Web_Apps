# App.py
# -----------------------------------------------------------------------------
# This file is the Flask web app wrapper for the Carasolva user creation workflow.
#
# It provides the browser-based page where a user can enter Carasolva credentials,
# upload a spreadsheet, choose the role to assign, and provide the Microsoft Edge
# WebDriver path. The app then starts the Selenium automation script in the
# background, streams live output to the browser with Server-Sent Events (SSE),
# and stores final run status/summary data in memory for the front end to display.
#
# SECURITY / PUBLIC REPO NOTE:
# This copy uses dummy placeholder values for paths, logo names, default roles,
# and any environment-specific values. Replace them with real values only in a
# private/internal deployment environment.
# -----------------------------------------------------------------------------

from flask import Flask, render_template, request, Response, abort, jsonify, url_for
import threading
import queue
import time
import os
import sys
import subprocess
from collections import defaultdict
from datetime import datetime
import traceback
import pandas as pd
from uuid import uuid4

# --------- Paths & folders ----------
BASE_DIR = os.path.dirname(__file__)
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
STATIC_DIR = os.path.join(BASE_DIR, "static")
UPLOAD_ROOT = os.path.join(BASE_DIR, "uploads")
os.makedirs(STATIC_DIR, exist_ok=True)
os.makedirs(UPLOAD_ROOT, exist_ok=True)

# --------- Flask setup ----------
# Serve static files from /static (standard Flask behavior)
app = Flask(__name__, template_folder=TEMPLATES_DIR, static_folder=STATIC_DIR, static_url_path="/static")
app.config['MAX_CONTENT_LENGTH'] = 64 * 1024 * 1024  # 64MB upload cap

# Default Edge driver path placeholder (user can override in the form).
DEFAULT_DRIVER_PATH = r"C:\Path\To\msedgedriver.exe"

# Path to your Carasolva script (same folder as app.py by default)
# NOTE: matches your actual filename with a space.
CARASOLVA_SCRIPT = os.path.join(BASE_DIR, "Carasolva UserCreation.py")

# --------- SSE multi-client per run_id ----------
user_channels = defaultdict(set)  # run_id -> set[Queue]
channels_lock = threading.Lock()

# Final results per run_id (what the UI renders at the end)
# { run_id: { "status": "running|done|error", "finished_at": "iso", "all_users": [...], "workdir": "..." } }
run_results = {}

# Sends one log/status message to every browser client listening to the same run_id.
# This lets multiple open tabs receive the same live automation output.
def broadcast(run_id: str, message: str):
    with channels_lock:
        chans = list(user_channels.get(run_id, set()))
    for q in chans:
        try:
            q.put_nowait(message)
        except Exception:
            pass

# Creates and registers a queue for one browser EventSource connection.
# Each connected client gets its own queue so live messages can be delivered independently.
def register_stream(run_id: str) -> queue.Queue:
    q = queue.Queue()
    with channels_lock:
        user_channels[run_id].add(q)
    return q

# Removes a browser client's queue when the EventSource connection closes.
# It also cleans up the run_id entry if no clients are still listening.
def unregister_stream(run_id: str, q: queue.Queue):
    with channels_lock:
        if run_id in user_channels and q in user_channels[run_id]:
            user_channels[run_id].remove(q)
        if run_id in user_channels and not user_channels[run_id]:
            user_channels.pop(run_id, None)

# --------- Helpers ----------
ALLOWED_EXTS = {".xlsx", ".xls", ".csv"}

# Checks whether the uploaded spreadsheet file has an allowed extension.
# This prevents unsupported file types from being accepted by the web app.
def safe_ext(filename: str) -> bool:
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXTS

# Creates a separate upload/work folder for a specific automation run.
# Keeping each run in its own folder prevents uploaded spreadsheets from overwriting each other.
def ensure_user_workdir(run_id: str) -> str:
    workdir = os.path.join(UPLOAD_ROOT, run_id)
    os.makedirs(workdir, exist_ok=True)
    return workdir

# Reads the uploaded spreadsheet and extracts basic user info for the final web summary.
# This does not create users; it only prepares a lightweight name/email list for display.
def read_users_table_generic(path):
    """Read .xlsx/.xls/.csv and produce [{'name':..., 'email':...}, ...] for summary."""
    ext = os.path.splitext(path)[1].lower()
    if ext in [".xlsx", ".xls"]:
        df = pd.read_excel(path)
    elif ext == ".csv":
        df = pd.read_csv(path)
    else:
        return []

    # normalize headers
    norm_map = {}
    for c in df.columns:
        key = (
            str(c).strip().lower()
            .replace("_", " ").replace("-", " ").replace(".", " ")
        )
        key = " ".join(key.split())
        norm_map[c] = key
    df = df.rename(columns=norm_map)

    FIRST_CANDS = {"first name", "firstname", "first", "f name", "given name"}
    LAST_CANDS  = {"last name", "lastname", "last", "l name", "surname", "family name"}
    EMAIL_CANDS = {"email", "e-mail", "email address", "mail"}

    def choose(colset):
        for c in df.columns:
            if c in colset:
                return c
        return None

    col_first = choose(FIRST_CANDS)
    col_last  = choose(LAST_CANDS)
    col_email = choose(EMAIL_CANDS)

    users = []
    for _, r in df.iterrows():
        first = str(r.get(col_first, "")).strip() if col_first else ""
        last  = str(r.get(col_last, "")).strip() if col_last else ""
        email = str(r.get(col_email, "")).strip() if col_email else ""
        name = f"{first} {last}".strip()
        if name or email:
            users.append({"name": name, "email": email})
    return users

# Reads the Selenium subprocess output line by line.
# Every line from the automation script is broadcast to the browser as live status output.
def stream_process_stdout(proc, run_id):
    """Read stdout of subprocess line by line and broadcast to the browser."""
    try:
        for raw in iter(proc.stdout.readline, b''):
            if not raw:
                break
            try:
                line = raw.decode(errors="ignore").rstrip("\n")
            except Exception:
                line = raw.decode("utf-8", errors="ignore").rstrip("\n")
            if line.strip():
                broadcast(run_id, line)
    except Exception as e:
        broadcast(run_id, f"⚠️ Stream error: {e}")

# Runs the Carasolva Selenium script in a background worker thread.
# This keeps the Flask web request from freezing while the browser automation is running.
def run_carasolva_in_worker(run_id, username, password, file_path, role_text, driver_path):
    """Worker that runs the external Carasolva script and streams stdout to SSE."""
    try:
        broadcast(run_id, "Starting browser and navigating to login page...")

        # Build command (use list to handle spaces in path)
        cmd = [
            sys.executable, CARASOLVA_SCRIPT,
            username, password,
            "--file", file_path,
            "--role", role_text,
            "--driver", driver_path
        ]

        # Spawn subprocess
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            bufsize=1
        )

        # Stream output
        stream_process_stdout(proc, run_id)

        # Wait until finished
        rc = proc.wait()

        if rc == 0:
            run_results[run_id]["status"] = "done"
            run_results[run_id]["finished_at"] = datetime.utcnow().isoformat() + "Z"
            broadcast(run_id, "Script finished. Preparing summary…")
        else:
            run_results[run_id]["status"] = "error"
            run_results[run_id]["finished_at"] = datetime.utcnow().isoformat() + "Z"
            broadcast(run_id, f"❌ Script exited with code {rc}. See logs above.")

    except Exception as e:
        run_results[run_id]["status"] = "error"
        run_results[run_id]["finished_at"] = datetime.utcnow().isoformat() + "Z"
        broadcast(run_id, f"❌ Fatal error: {e}\n{traceback.format_exc()}")
    finally:
        # Tell clients to close their EventSource
        broadcast(run_id, "[DONE]")

# --------- Routes ----------
# Renders the main Carasolva web form.
# The page lets the user enter credentials, upload a spreadsheet, choose a role, and start the job.
@app.route("/")
def index():
    # Optional: prints the placeholder logo path so template/static issues are easy to spot in console.
    logo_url = url_for('static', filename='company-logo-placeholder.png')
    print(f"[INFO] Logo expected at: {logo_url}")
    return render_template("index.html")

# Handles the form submission from the browser.
# It validates inputs, saves the uploaded spreadsheet, creates a run_id, and starts the worker thread.
@app.route("/start", methods=["POST"])
def start():
    username = request.form.get("username", "").strip()
    password = request.form.get("password", "").strip()
    role_text = request.form.get("role", "").strip() or "Example Role"
    driver_path = request.form.get("driver_path", "").strip() or DEFAULT_DRIVER_PATH

    if not username or not password:
        return "Missing username or password", 400

    if "file" not in request.files:
        return "Missing spreadsheet", 400

    f = request.files["file"]
    if not f or not f.filename:
        return "Missing spreadsheet filename", 400

    if not safe_ext(f.filename):
        return "Unsupported file type. Upload .xlsx / .xls / .csv", 400

    # Unique run id
    run_id = f"{username}-{uuid4().hex[:8]}"

    # Per-run workdir
    workdir = ensure_user_workdir(run_id)
    save_path = os.path.join(workdir, f.filename)
    f.save(save_path)

    # Pre-read for summary (even if SSE hiccups)
    try:
        all_users = read_users_table_generic(save_path)
    except Exception as e:
        all_users = []
        print("Failed to read users for summary:", e)

    run_results[run_id] = {
        "status": "running",
        "finished_at": None,
        "all_users": all_users,
        "workdir": workdir
    }

    # Kick off background worker
    t = threading.Thread(
        target=run_carasolva_in_worker,
        args=(run_id, username, password, save_path, role_text, driver_path),
        daemon=True
    )
    t.start()

    # Return run_id so the client can subscribe/poll this job
    return jsonify({"run_id": run_id}), 200

# Opens the Server-Sent Events stream for one automation run.
# The browser connects here to receive real-time log messages from the background worker.
@app.route("/stream")
def stream():
    run_id = request.args.get("run_id", "").strip()
    if not run_id:
        abort(400)

    client_q = register_stream(run_id)

    def event_stream():
        try:
            # Make EventSource auto-retry quicker
            yield "retry: 1500\n\n"
            while True:
                try:
                    line = client_q.get(timeout=20)
                    print(f"Streaming line ({run_id}): {line}")
                    yield f"data: {line}\n\n"
                    if line == "[DONE]":
                        # small keep-alive burst before exit
                        for _ in range(2):
                            yield ": keep-alive\n\n"
                            time.sleep(0.2)
                        break
                except queue.Empty:
                    yield ": keep-alive\n\n"
        finally:
            unregister_stream(run_id, client_q)

    headers = {
        "Cache-Control": "no-cache",
        "X-Accel-Buffering": "no",
    }
    return Response(event_stream(), headers=headers, mimetype="text/event-stream")

# Returns the final status and user summary for a completed or running job.
# The front end polls this endpoint so it can show the final result even if the SSE stream hiccups.
@app.route("/result")
def result():
    run_id = request.args.get("run_id", "").strip()
    if not run_id:
        abort(400)
    data = run_results.get(run_id)
    if not data:
        return jsonify({"status": "unknown"}), 404
    return jsonify(data)

# Simple health-check endpoint.
# This can be used to confirm the Flask app is running.
@app.route("/health")
def health():
    return "ok", 200

if __name__ == "__main__":
    # threaded=True helps SSE + worker coexist; debug=True for dev
    # Use host 0.0.0.0 so others on LAN can access it.
    app.run(host="0.0.0.0", port=5000, debug=True, threaded=True)
