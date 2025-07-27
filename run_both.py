import subprocess
import os
import time
from pathlib import Path

# --- Paths ---
script_dir = os.path.dirname(os.path.abspath(__file__))
qtm_path = os.path.join(script_dir, "qtm_log.py")
fitts_path = os.path.join(script_dir, "fitts.py")
session_path_file = os.path.join(script_dir, "current_session_path.txt")

# --- Step 1: Start Fitts experiment in a separate process ---
fitts_proc = subprocess.Popen(["python", fitts_path])

# --- Step 2: Wait for the current_session_path.txt file to appear ---
timeout = 30  # Allow longer time
waited = 0
while not os.path.exists(session_path_file) and waited < timeout:
    time.sleep(0.5)
    waited += 0.5

if not os.path.exists(session_path_file):
    print("❌ Session path not found in time.")
    fitts_proc.terminate()
    exit(1)

# --- Step 3: Read the session folder path ---
with open(session_path_file, "r") as f:
    folder_path = f.read().strip()

# --- Step 4: Start QTM logger in parallel, passing the folder path ---
print(f"📁 Logging QTM data to: {folder_path}")
qtm_proc = subprocess.Popen(["python", qtm_path, folder_path])

# --- Step 5: Wait for the Fitts experiment to end ---
fitts_proc.wait()

# --- Step 6: Then kill the QTM logger and exit ---
print("🛑 Experiment finished. Closing QTM logger...")
qtm_proc.terminate()
print("✅ All done. QTM logger and experiment closed.")
