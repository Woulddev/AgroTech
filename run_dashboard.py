#!/usr/bin/env python3
"""
PRIVA VP9508 Greenhouse Sensor Dashboard
Run: python3 run_dashboard.py
Then open: http://localhost:5000
"""
import importlib.util
import subprocess
import sys
import os

print("📦 Checking dependencies...")
required = {
    "flask": "flask",
    "openpyxl": "openpyxl",
}
missing = [package for module, package in required.items() if importlib.util.find_spec(module) is None]
if missing:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])

# Run the server
os.chdir(os.path.dirname(os.path.abspath(__file__)))
subprocess.run([sys.executable, "app.py"])
