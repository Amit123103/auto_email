"""
Skytecher Cold Email Automation - Application Launcher
======================================================
Entrypoint that runs the unified Streamlit dashboard (main.py).
You can start the app using either:
    streamlit run main.py
    python main.py
    python app.py
"""

import os
import sys
import subprocess
from pathlib import Path

# Ensure UTF-8 console encoding on Windows
os.environ["PYTHONIOENCODING"] = "utf-8"
os.environ["PYTHONUTF8"] = "1"

BASE_DIR = Path(__file__).resolve().parent
MAIN_FILE = BASE_DIR / "main.py"


def run_app():
    """Launch the main Streamlit application dashboard on port 8501."""
    cmd = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        str(MAIN_FILE),
        "--server.port",
        "8501",
    ] + sys.argv[1:]
    try:
        sys.exit(subprocess.call(cmd))
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    run_app()
