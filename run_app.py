import contextlib
import http.client
import os
import socket
import subprocess
import sys
import time
import traceback
import webbrowser


APP_VARIANT = os.environ.get("PCS_APP_ENV", "production").strip().lower() or "production"
APP_DISPLAY_NAME = os.environ.get("PCS_APP_DISPLAY_NAME", "PCS Proposal").strip() or "PCS Proposal"
DEFAULT_PORT = int(os.environ.get("PCS_DEFAULT_PORT", "5050"))
APP_STATE_DIR = os.environ.get("PCS_APP_STATE_DIR", "/tmp/pcs_proposal_app")
SERVER_STARTUP_TIMEOUT = float(os.environ.get("PCS_SERVER_STARTUP_TIMEOUT", "30"))
APP_STARTUP_LOG_PATH = os.path.join(APP_STATE_DIR, "startup.log")
SERVER_ONLY_ENV = "PCS_PROPOSAL_SERVER_ONLY"
DESKTOP_LIFECYCLE_ENV = "PCS_PROPOSAL_DESKTOP_LIFECYCLE"
LOCAL_ROOF_WORKER_ENV = "ROOF_INTELLIGENCE_LOCAL_WORKER"


def _log_startup(message):
    line = f"{time.strftime('%Y-%m-%d %H:%M:%S')} {message}"
    try:
        print(line, flush=True)
    except Exception:
        pass
    try:
        os.makedirs(APP_STATE_DIR, exist_ok=True)
        with open(APP_STARTUP_LOG_PATH, "a", encoding="utf-8") as handle:
            handle.write(f"{line}\n")
    except Exception:
        pass


def _pick_port():
    requested = os.environ.get("PORT")
    if requested:
        port = int(requested)
        _log_startup(f"Using PORT environment override: {port}")
        return port

    with contextlib.closing(socket.socket(socket.AF_INET, socket.SOCK_STREAM)) as sock:
        try:
            sock.bind(("127.0.0.1", DEFAULT_PORT))
            _log_startup(f"Using default port {DEFAULT_PORT}")
            return DEFAULT_PORT
        except OSError:
            _log_startup(f"Default port {DEFAULT_PORT} is unavailable; selecting an ephemeral port")

    with contextlib.closing(socket.socket(socket.AF_INET, socket.SOCK_STREAM)) as sock:
        sock.bind(("127.0.0.1", 0))
        port = sock.getsockname()[1]
        _log_startup(f"Using ephemeral port {port}")
        return port


def _open_with_mac_open(app_name, url):
    try:
        cmd = ["open", "-a", app_name]
        if url:
            cmd.append(url)
        subprocess.run(
            cmd,
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=5,
        )
        return True
    except Exception as exc:
        _log_startup(f"open -a {app_name} failed: {exc}")
        return False


def _wait_for_server(host, port, timeout=30.0):
    deadline = time.time() + timeout
    last_error = None
    while time.time() < deadline:
        conn = None
        try:
            conn = http.client.HTTPConnection(host, port, timeout=0.5)
            conn.request("GET", "/")
            response = conn.getresponse()
            response.read()
            _log_startup(f"Server responded on http://{host}:{port}/ with status {response.status}")
            return True
        except Exception as exc:
            last_error = exc
            time.sleep(0.2)
        finally:
            if conn is not None:
                try:
                    conn.close()
                except Exception:
                    pass

    _log_startup(f"Timed out waiting for server on http://{host}:{port}/: {last_error}")
    return False


def _open_app_url(url):
    if sys.platform == "darwin" and _open_with_mac_open("Safari", url):
        _log_startup("Opened app URL in Safari")
        return True

    try:
        opened = bool(webbrowser.open_new(url))
        _log_startup(f"webbrowser.open_new returned {opened}")
        return opened
    except Exception as exc:
        _log_startup(f"webbrowser.open_new failed: {exc}")
        return False


def _run_server(host, port):
    _log_startup(f"Server process starting at http://{host}:{port}")
    from pcs_proposal_web import app

    app.run(host=host, port=port, debug=False, use_reloader=False)
    _log_startup(f"Server process stopped at http://{host}:{port}")


def _start_server_process(host, port):
    env = os.environ.copy()
    env[SERVER_ONLY_ENV] = "1"
    env[DESKTOP_LIFECYCLE_ENV] = "1"
    env.setdefault(LOCAL_ROOF_WORKER_ENV, "1")
    env["PORT"] = str(port)
    env["PCS_PROPOSAL_HOST"] = host

    os.makedirs(APP_STATE_DIR, exist_ok=True)
    log_handle = open(APP_STARTUP_LOG_PATH, "a", encoding="utf-8")
    try:
        process = subprocess.Popen(
            [sys.executable],
            stdin=subprocess.DEVNULL,
            stdout=log_handle,
            stderr=log_handle,
            env=env,
            close_fds=True,
            start_new_session=True,
        )
    finally:
        log_handle.close()

    _log_startup(f"Started detached server process pid {process.pid}")
    return process


def main():
    host = "127.0.0.1"
    _log_startup(f"{APP_DISPLAY_NAME} app launcher starting ({APP_VARIANT})")

    try:
        if os.environ.get(SERVER_ONLY_ENV) == "1":
            server_host = os.environ.get("PCS_PROPOSAL_HOST", host)
            server_port = int(os.environ.get("PORT", DEFAULT_PORT))
            _run_server(server_host, server_port)
            return

        port = _pick_port()
        url = f"http://{host}:{port}"
        process = _start_server_process(host, port)

        _log_startup(f"Waiting for Flask server at {url}")
        if not _wait_for_server(host, port, timeout=SERVER_STARTUP_TIMEOUT):
            if process.poll() is not None:
                _log_startup(f"Server process exited early with code {process.returncode}")
            return

        for attempt in range(5):
            _log_startup(f"Opening browser attempt {attempt + 1} for {url}")
            if _open_app_url(url):
                return
            time.sleep(0.5 + (attempt * 0.5))
        _log_startup(f"Unable to open browser after 5 attempts: {url}")
    except Exception:
        _log_startup("Fatal startup error:")
        _log_startup(traceback.format_exc())
        raise


if __name__ == "__main__":
    main()
