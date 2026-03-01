"""Launcher that sets CWD before importing streamlit to avoid PermissionError."""
import os
os.chdir("/Users/diegogaldeano/Desktop/claude/audit-app")

import sys
sys.argv = ["streamlit", "run", "/Users/diegogaldeano/Desktop/claude/audit-app/app.py",
            "--server.port", "8501", "--server.headless", "true"]

from streamlit.web.cli import main
main()
