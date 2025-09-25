# run_app.py
import os
import threading
import time
import webview

# --- Set defaults for your folders (override via env if needed) ---
os.environ.setdefault("PROPOSALS_DIR", "/Users/vernabbott/OneDrive/Professional Coating Systems/Proposals")
os.environ.setdefault("CONTRACTS_DIR", "/Users/vernabbott/OneDrive/Professional Coating Systems/Contracts")
os.environ.setdefault("COMPLETED_DIR", "/Users/vernabbott/OneDrive/Professional Coating Systems/Completed")
os.environ.setdefault("DEADFILE_DIR", "/Users/vernabbott/OneDrive/Professional Coating Systems/Dead Files")

# Optional: pin host/port
HOST = os.environ.get("PCS_HOST", "127.0.0.1")
PORT = int(os.environ.get("PCS_PORT", "5000"))

# Import your Flask app
from pcs_proposal_web import app  # assumes your app = Flask(__name__)

def run_flask():
    # threaded=True so it can serve while pywebview runs
    app.run(host=HOST, port=PORT, debug=False, threaded=True)

if __name__ == "__main__":
    t = threading.Thread(target=run_flask, daemon=True)
    t.start()
    # Give the server a moment to come up
    time.sleep(0.75)

    # Open native window to your app
    url = f"http://{HOST}:{PORT}/"
    window = webview.create_window("PCS Proposals", url, width=1280, height=800, resizable=True)
    # When window closes, the process exits (flask thread is daemon)
    webview.start()