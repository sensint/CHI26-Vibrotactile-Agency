import os
import time
import csv
import threading
import tkinter as tk
from tkinter import simpledialog, messagebox
from screeninfo import get_monitors
import ctypes
from pathlib import Path
import sys
import zmq
import json

# ZeroMQ Configuration
ZMQ_HOST = "localhost"
ZMQ_PORT = 5555
ZMQ_CONFIG_PORT = 5556  # Port for receiving config from publisher
ZMQ_TOPIC = "qtm_data"

# Default W/D values (will be overridden by publisher config)
W_VALUES = [80]
D_VALUES = [240]
TOTAL_TRIALS = 10

# These will be received from the publisher via ZMQ
participant_folder = ""
participant_name = ""
conditions = ""
attempts = 0
ID = 0
delaytime = 650

participant_id = ""
trial_count = {}
difficulty = 1
clicks = 0
target_side = 1
previous_rects = []
data = []
target_sides = []
current_rect = (0, 0, 0, 0)
start_time = 0
CANVAS_WIDTH = 0
CANVAS_HEIGHT = 0
experiment_finished = False

# ZMQ State Machine Variables
latest_frame = None
previous_inside = False
trigger_detected = False

# Global variables for target bounds in mm
rect_x_mm = None
rect_x_end_mm = None

script_dir = Path(__file__).parent
session_file = script_dir / "current_session_path.txt"

monitors = get_monitors()
selected_monitor = monitors[1] if len(monitors) > 1 else monitors[0]
root = tk.Tk()
root.withdraw()
experiment_window = tk.Toplevel()
screen_x = selected_monitor.x
screen_y = selected_monitor.y
screen_width = selected_monitor.width
screen_height = selected_monitor.height
experiment_window.geometry(f"{screen_width}x{screen_height}+{screen_x}+{screen_y}")
experiment_window.update_idletasks()
hwnd = ctypes.windll.user32.GetForegroundWindow()
ctypes.windll.user32.MoveWindow(hwnd, screen_x, screen_y, screen_width, screen_height, True)
CANVAS_WIDTH = screen_width
CANVAS_HEIGHT = screen_height
canvas = tk.Canvas(experiment_window, width=CANVAS_WIDTH, height=CANVAS_HEIGHT, bg="white")
canvas.pack()
canvas.pack_forget()

def listen_zmq():
    """ZMQ listener with state machine trigger detection."""
    global latest_frame, previous_inside, trigger_detected, current_rect, rect_x_mm, rect_x_end_mm
    
    ctx = zmq.Context()
    socket = ctx.socket(zmq.SUB)
    
    connect_addr = f"tcp://{ZMQ_HOST}:{ZMQ_PORT}"
    socket.connect(connect_addr)
    socket.setsockopt_string(zmq.SUBSCRIBE, ZMQ_TOPIC)
    
    print(f"\U0001F4E1 Connected to {connect_addr}")
    print(f"\U0001F4E5 Subscribed to topic: '{ZMQ_TOPIC}'")
    print("\U0001F7E2 Waiting for data...\n")

    while not experiment_finished:
        try:
            message = socket.recv_string()
            _, json_str = message.split(" ", 1)
            data = json.loads(json_str)
            
            frame = data['frame']
            x_local = data.get('x_local')
            status = data.get('status', 'unknown')
            latest_frame = frame
            
            # Calculate inside_bounds based on x_local position and actual target bounds
            # 1 = inside bounds, 0 = outside bounds
            if x_local is not None and rect_x_mm is not None and rect_x_end_mm is not None:
                # Convert x_local to absolute value for comparison
                x_local_abs = abs(x_local)
                inside_bounds = (rect_x_mm <= x_local_abs <= rect_x_end_mm)
            else:
                inside_bounds = False
            
            state_value = 1 if inside_bounds else 0

            # Build rect_bounds_str using global variables
            if rect_x_mm is not None and rect_x_end_mm is not None:
                rect_bounds_str = f" | Target bounds (mm): {rect_x_mm:.2f} to {rect_x_end_mm:.2f}"
            else:
                rect_bounds_str = ""

            print(f"Frame {frame:6d} | State: {state_value} | X_local: {x_local} | Status: {status} | Prev: {1 if previous_inside else 0}{rect_bounds_str}", end="")
            
            # State Machine: Trigger only on transition from 0 → 1 (outside → inside)
            if inside_bounds and not previous_inside:
                # Transition detected: 0 → 1
                trigger_detected = True
                experiment_window.event_generate('<<ZMQTrigger>>', when='tail')
                print(" | \U0001F518 TRIGGER DETECTED!")
            else:
                print()
            
            # Update state for next iteration
            previous_inside = inside_bounds
            
        except Exception as e:
            print(f"❌ ZMQ error: {e}")
            time.sleep(0.1)


    socket.close()
    ctx.term()
    print("🛑 ZMQ listener stopped.")

def get_rect_bounds_str():
    global current_rect, screen_width_mm

    rect_x, _, rect_w, _ = current_rect

    if rect_w == 0:
        return ""

    mm_per_px_x = screen_width_mm / CANVAS_WIDTH
    left_mm = rect_x * mm_per_px_x
    right_mm = (rect_x + rect_w) * mm_per_px_x

    return f" | Target(mm): {left_mm:.2f}->{right_mm:.2f}"


def draw_rectangle():
    global current_rect, start_time, rect_x_mm, rect_x_end_mm
    canvas.delete("all")
    for rect in previous_rects:
        canvas.create_rectangle(rect[0], rect[1], rect[0]+rect[2], rect[1]+rect[3], fill="lightgrey")
    rect_width = W_VALUES[difficulty - 1]
    distance = D_VALUES[difficulty - 1]
    rect_height = CANVAS_HEIGHT
    center_x = CANVAS_WIDTH / 2
    rect_x = center_x + (target_side * (distance / 2)) - (rect_width / 2)
    rect_y = 0
    canvas.create_rectangle(rect_x, rect_y, rect_x + rect_width, rect_y + rect_height, fill="blue")
    current_rect = (rect_x, rect_y, rect_width, rect_height)
    start_time = time.time()
    
    # Calculate and store target bounds in mm as global variables
    screen_width_mm, _ = get_screen_dimensions_mm()
    px_to_mm = screen_width_mm / CANVAS_WIDTH if CANVAS_WIDTH else 1
    rect_x_mm = rect_x * px_to_mm
    rect_x_end_mm = (rect_x + rect_width) * px_to_mm
    
    print(f"\nTarget bounds (mm): {rect_x_mm:.2f} to {rect_x_end_mm:.2f}\n")

def get_screen_dimensions_mm():
    try:
        files = sorted(Path(participant_folder).glob("clicked_log_*.csv"), reverse=True)
        if not files:
            return 346.00, 194.57
        with open(files[0], 'r', newline='') as f:
            reader = csv.reader(f)
            next(reader)  # skip header
            x_vals = []
            y_vals = []
            for r in reader:
                if len(r) > 6 and r[5] != 'NaN' and r[6] != 'NaN':
                    x_vals.append(abs(float(r[5])))
                    y_vals.append(abs(float(r[6])))
        return max(x_vals)*2 if x_vals else 346.0, max(y_vals)*2 if y_vals else 194.57
    except:
        return 346.0, 194.57

def handle_zmq_trigger(event=None):
    """Handle trigger from ZMQ instead of serial."""
    global clicks, target_side, current_rect, experiment_finished, trigger_detected
    
    if experiment_finished:
        return
    
    # Mark frame as clicked in qtm_TB.py's clicked_frames
    from qtm_TB import clicked_frames
    if latest_frame is not None:
        clicked_frames.add(latest_frame)
    
    click_time = time.time()
    mt = (click_time - start_time) * 1000
    speed = D_VALUES[difficulty - 1] / mt if mt > 0 else 0
    throughput = difficulty / (mt / 1000) if mt > 0 else 0
    data.append({'MT': mt, 'speed': speed, 'throughput': throughput})
    target_sides.append(target_side)
    previous_rects.append(current_rect)
    clicks += 1
    target_side *= -1
    
    trigger_detected = False
    
    if clicks >= TOTAL_TRIALS:
        experiment_finished = True
        experiment_window.unbind('<<ZMQTrigger>>')
        end_trial()
    else:
        draw_rectangle()

def save_data_and_finish():
    global experiment_finished, participant_folder  
    print(f"� {TOTAL_TRIALS} triggers done. Saving subscriber data...")
    # Exclude first 3 trials from averages (first trials are positioning, not real Fitts' movements)
    valid_data = data[3:] if len(data) > 3 else data
    avg_mt = sum(d['MT'] for d in valid_data) / len(valid_data)
    avg_speed = sum(d['speed'] for d in valid_data) / len(valid_data)
    avg_tp = sum(d['throughput'] for d in valid_data) / len(valid_data)

    headers = ['MT', 'speed', 'throughput', 'LocalX', 'participant_name', 'conditions', 'ID', 'attempt', 'delaytime']

    filename = f"{participant_name}_{conditions}_ID{ID}_{attempts}_{delaytime}.csv"
    filepath = os.path.join(participant_folder, filename)

    with open(filepath, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        for row in data:
            writer.writerow([
                row.get('MT', ''),
                row.get('speed', ''),
                row.get('throughput', ''),
                "",
                participant_name,
                conditions,
                ID,
                attempts,
                delaytime
            ])
        # Averages row (exclude first 3 trials)
        writer.writerow([
            f"{avg_mt:.4f}",
            f"{avg_speed:.6f}",
            f"{avg_tp:.4f}",
        ])
    print(f"✅ Data saved to: {filepath}")

    # Wait for publisher's clicked_log (with timeout)
    clicked_log = os.path.join(participant_folder, f"{participant_name}_{conditions}_ID{ID}_{attempts}_{delaytime}_clicked_log.csv")
    for _ in range(20):  # Wait up to 10 seconds
        if os.path.exists(clicked_log):
            break
        time.sleep(0.5)

    summary = f"Avg MT: {avg_mt:.2f} ms\nAvg Speed: {avg_speed:.2f} px/ms\nAvg Throughput: {avg_tp:.2f} bit/s"
    print(summary)

STAY_RED_MS = 3_000

def end_trial():
    global current_rect
    rect_x, rect_y, rect_width, rect_height = current_rect

    def show_red():
        canvas.delete("all")
        canvas.create_rectangle(rect_x, rect_y, rect_x + rect_width, rect_y + rect_height, fill="red")
        experiment_window.update()

        def quit_after_red():
            save_data_and_finish()
            print("👋 Subscriber exiting...")
            os._exit(0)

        experiment_window.after(STAY_RED_MS, quit_after_red)

    experiment_window.after(delaytime, show_red)

def begin_trial():
    global difficulty, clicks, data, previous_rects, target_sides
    difficulty = 1
    clicks = 0
    data.clear()
    target_sides.clear()
    previous_rects.clear()
    trial_count.setdefault(participant_id, {}).setdefault(difficulty, 0)
    trial_count[participant_id][difficulty] += 1
    canvas.pack()
    draw_rectangle()

def wait_for_config():
    """Wait for config from publisher via ZMQ config channel."""
    global participant_name, participant_folder, conditions, attempts, ID, delaytime
    global W_VALUES, D_VALUES, TOTAL_TRIALS

    ctx = zmq.Context()
    socket = ctx.socket(zmq.SUB)
    socket.connect(f"tcp://{ZMQ_HOST}:{ZMQ_CONFIG_PORT}")
    socket.setsockopt_string(zmq.SUBSCRIBE, "config")
    print("⏳ Waiting for config from publisher...")

    message = socket.recv_string()
    _, json_str = message.split(" ", 1)
    config = json.loads(json_str)

    participant_name = config["participant_name"]
    participant_folder = config["participant_folder"]
    conditions = config["conditions"]
    attempts = config["attempts"]
    ID = config["ID"]
    delaytime = config["delaytime"]
    W_VALUES = config.get("W_VALUES", W_VALUES)
    D_VALUES = config.get("D_VALUES", D_VALUES)
    TOTAL_TRIALS = config.get("TOTAL_TRIALS", TOTAL_TRIALS)

    socket.close()
    ctx.term()
    print(f"✅ Config received: {participant_name}, {conditions}, ID{ID}, attempt {attempts}, delay {delaytime}")


def start_experiment():
    global participant_name, participant_folder

    os.makedirs(participant_folder, exist_ok=True)
    session_file.write_text(participant_folder)
    begin_trial()

# Start button (shown after config is received)
btn_start = tk.Button(experiment_window, text="Start Experiment", font=("Arial", 20),
                      command=lambda: [btn_start.pack_forget(), start_experiment()])
btn_start.pack(pady=10)
btn_start.config(state=tk.DISABLED)  # Disabled until config arrives

# Wait for config from publisher, then enable Start button
def on_config_received():
    wait_for_config()
    experiment_window.after(0, lambda: btn_start.config(state=tk.NORMAL))

# Start config listener in background thread
threading.Thread(target=on_config_received, daemon=True).start()

# Start ZMQ data listener thread
threading.Thread(target=listen_zmq, daemon=True).start()

# Bind ZMQ trigger instead of serial
experiment_window.bind('<<ZMQTrigger>>', handle_zmq_trigger)
experiment_window.mainloop()