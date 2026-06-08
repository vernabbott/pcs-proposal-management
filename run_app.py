import contextlib
import http.client
import os
import socket
import threading
import time
import subprocess
import sys
import traceback
import webbrowser


def _log_startup(message):
    try:
        print(message, flush=True)
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
            sock.bind(("127.0.0.1", 5050))
            _log_startup("Using default port 5050")
            return 5050
        except OSError:
            _log_startup("Default port 5050 is unavailable; selecting an ephemeral port")
            pass

    with contextlib.closing(socket.socket(socket.AF_INET, socket.SOCK_STREAM)) as sock:
        sock.bind(("127.0.0.1", 0))
        port = sock.getsockname()[1]
        _log_startup(f"Using ephemeral port {port}")
        return port


def _run_osascript(script_lines):
    try:
        subprocess.run(
            ["osascript", *sum((["-e", line] for line in script_lines), [])],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        return True
    except Exception:
        return False


def _is_process_running(name):
    try:
        result = subprocess.run(
            ["pgrep", "-x", name],
            check=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        return result.returncode == 0
    except Exception:
        return False


def _open_with_mac_open(app_name, url):
    try:
        cmd = ["open", "-a", app_name]
        if url:
            cmd.append(url)
        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except Exception as exc:
        _log_startup(f"open -a {app_name} failed: {exc}")
        return False


def _escape_applescript_string(value):
    return value.replace("\\", "\\\\").replace('"', '\\"')


def _open_in_safari(url):
    if sys.platform != "darwin":
        return False

    if not _is_process_running("Safari"):
        return _open_with_mac_open("Safari", url)

    escaped_url = _escape_applescript_string(url)
    script_lines = [
        'tell application "Safari"',
        f'set targetURL to "{escaped_url}"',
        "activate",
        "repeat with safariWindow in windows",
        "repeat with safariTab in tabs of safariWindow",
        'set tabURL to URL of safariTab',
        'set tabName to name of safariTab',
        'if (tabURL starts with "http://127.0.0.1:") and ((tabName contains "PCS") or (tabName contains "Proposal")) then',
        "set current tab of safariWindow to safariTab",
        "set index of safariWindow to 1",
        "set URL of safariTab to targetURL",
        "return",
        "end if",
        "end repeat",
        "end repeat",
        "if (count of windows) is 0 then",
        "make new document with properties {URL:targetURL}",
        "else",
        "tell front window to set current tab to (make new tab with properties {URL:targetURL})",
        "end if",
        "end tell",
    ]

    if _run_osascript(script_lines):
        return True

    return _open_with_mac_open("Safari", url)


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


def _is_existing_pcs_server(host, port, timeout=0.8):
    conn = None
    try:
        conn = http.client.HTTPConnection(host, port, timeout=timeout)
        conn.request("GET", "/")
        response = conn.getresponse()
        body = response.read(20000).decode("utf-8", errors="ignore")
        return response.status < 500 and (
            "PCS Management" in body or "Proposal Management" in body
        )
    except Exception:
        return False
    finally:
        if conn is not None:
            try:
                conn.close()
            except Exception:
                pass


def _wait_for_port_to_close(host, port, timeout=2.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        with contextlib.closing(socket.socket(socket.AF_INET, socket.SOCK_STREAM)) as sock:
            sock.settimeout(0.2)
            try:
                sock.connect((host, port))
            except OSError:
                return True
        time.sleep(0.1)
    return False


def _open_app_url(url):
    if sys.platform == "darwin":
        if _open_in_safari(url):
            _log_startup("Opened app URL in Safari")
            return True
        if _open_with_mac_open("Safari", url):
            _log_startup("Opened app URL using macOS open")
            return True

    try:
        opened = bool(webbrowser.open_new(url))
        _log_startup(f"webbrowser.open_new returned {opened}")
        return opened
    except Exception as exc:
        _log_startup(f"webbrowser.open_new failed: {exc}")
        return False


def main():
    host = "127.0.0.1"
    _log_startup("PCS Proposal app launcher starting")

    try:
        if not os.environ.get("PORT") and _is_existing_pcs_server(host, 5050):
            existing_url = f"http://{host}:5050"
            _log_startup(f"Detected existing PCS server at {existing_url}; opening existing app")
            _open_app_url(existing_url)
            return

        from pcs_proposal_web import app

        port = _pick_port()
        url = f"http://{host}:{port}"

        def _open_browser():
            _log_startup(f"Waiting for Flask server at {url}")
            if not _wait_for_server(host, port):
                return
            for attempt in range(5):
                _log_startup(f"Opening browser attempt {attempt + 1} for {url}")
                if _open_app_url(url):
                    return
                time.sleep(0.5 + (attempt * 0.5))
            _log_startup(f"Unable to open browser after 5 attempts: {url}")

        threading.Thread(target=_open_browser, daemon=True).start()
        _log_startup(f"Starting Flask server at {url}")
        app.run(host=host, port=port, debug=False, use_reloader=False)
        _wait_for_port_to_close(host, port)
        _log_startup("Flask server stopped")
    except Exception:
        _log_startup("Fatal startup error:")
        _log_startup(traceback.format_exc())
        raise


if __name__ == "__main__":
    main()
