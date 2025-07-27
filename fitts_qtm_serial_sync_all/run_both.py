import subprocess
import os
import time
from pathlib import Path

script_dir = os.path.dirname(os.path.abspath(__file__))
fitts_path = os.path.join(script_dir, "fitts.py")
qtm_path = os.path.join(script_dir, "qtm_log.py")
session_file = Path(script_dir) / "current_session_path.txt"

# Step 1: Run Fitts' experiment (blocking)
subprocess.run(["python", fitts_path])

# Step 2: Wait for current_session_path.txt to exist
print("⏳ Waiting for participant folder info...")
timeout = 10  # seconds
start_time = time.time()
while not session_file.exists():
    if time.time() - start_time > timeout:
        print("❌ Session path not found in time.")
        exit(1)
    time.sleep(0.5)

with open(session_file, "r") as f:
    folder_path = f.read().strip()

if not folder_path:
    print("❌ Session path was empty.")
    exit(1)

full_path = os.path.join(script_dir, folder_path)
os.makedirs(full_path, exist_ok=True)

print(f"📁 Logging QTM to: {full_path}")

# Step 3: Start QTM logger and wait
qtm_proc = subprocess.Popen(["python", qtm_path, full_path])

try:
    print("📡 QTM logger is running. Press Ctrl+C to stop.")
    qtm_proc.wait()
except KeyboardInterrupt:
    print("🛑 Experiment interrupted. Stopping QTM logger.")
    qtm_proc.terminate()
