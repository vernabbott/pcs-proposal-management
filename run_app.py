# run_app.py
from pcs_proposal_web import app  # Flask app object

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)