import time
import threading
import tkinter as tk
from tkinter import simpledialog, messagebox
from openpyxl import Workbook
import os
import serial
from screeninfo import get_monitors
import ctypes
from collections import deque

SERIAL_PORT = 'COM7'
BAUD_RATE = 9600
W_VALUES = [90, 80, 70, 60, 50]
D_VALUES = [90, 160, 280, 480, 800]
TOTAL_TRIALS = 20

participant_name = ""
participant_id = ""
vibro_feedback = False
trial_count = {}
difficulty = 1
clicks = 0
target_side = 1
previous_rects = []
data = []
participant_folder = ""
current_rect = (0, 0, 0, 0)
start_time = 0
CANVAS_WIDTH = 0
CANVAS_HEIGHT = 0
experiment_finished = False

try:
    ser = serial.Serial(SERIAL_PORT, BAUD_RATE, timeout=0.1)
except:
    ser = None
    print("Warning: Serial port not available.")

monitors = get_monitors()
selected_monitor = monitors[1] if len(monitors) > 1 else monitors[0]
root = tk.Tk()
root.withdraw()
experiment_window = tk.Toplevel()
experiment_window.title("Fitts' Law Experiment")
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

def draw_rectangle():
    global current_rect, start_time
    canvas.delete("all")
    for rect in previous_rects:
        canvas.create_rectangle(rect[0], rect[1], rect[0] + rect[2], rect[1] + rect[3], fill="lightgrey")

    rect_width = W_VALUES[difficulty - 1]
    distance = D_VALUES[difficulty - 1]
    rect_height = CANVAS_HEIGHT
    center_x = CANVAS_WIDTH / 2
    rect_x = center_x + (target_side * distance) - (rect_width / 2)
    rect_y = 0
    canvas.create_rectangle(rect_x, rect_y, rect_x + rect_width, rect_y + rect_height, fill="blue")
    current_rect = (rect_x, rect_y, rect_width, rect_height)
    start_time = time.time()

def handle_serial_click(event=None):
    global clicks, target_side, current_rect, experiment_finished
    if experiment_finished:
        return

    click_time = time.time()
    mt = (click_time - start_time) * 1000
    error = False

    rect_x, rect_y, rect_w, rect_h = current_rect
    if not (rect_x <= CANVAS_WIDTH // 2 <= rect_x + rect_w and rect_y <= CANVAS_HEIGHT // 2 <= rect_y + rect_h):
        error = True

    speed = D_VALUES[difficulty - 1] / mt if mt > 0 else 0
    throughput = difficulty / (mt / 1000) if mt > 0 else 0

    data.append({
        'MT': mt,
        'speed': speed,
        'error': int(error),
        'throughput': throughput
    })

    previous_rects.append(current_rect)
    clicks += 1
    target_side *= -1

    if clicks >= TOTAL_TRIALS:
        experiment_finished = True
        end_trial()
    else:
        draw_rectangle()

def save_data_and_finish():
    global experiment_finished
    avg_mt = sum(d['MT'] for d in data) / TOTAL_TRIALS
    avg_speed = sum(d['speed'] for d in data) / TOTAL_TRIALS
    avg_tp = sum(d['throughput'] for d in data) / TOTAL_TRIALS
    avg_err = sum(d['error'] for d in data) / TOTAL_TRIALS
    data.append({
        'MT': avg_mt,
        'speed': avg_speed,
        'error': avg_err,
        'throughput': avg_tp
    })
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Results"
    headers = ['MT', 'speed', 'error', 'throughput']
    ws.append(headers)
    for row in data:
        ws.append([row.get(h, '') for h in headers])
    vibro_label = "VibroOn" if vibro_feedback else "VibroOff"
    filename = f"{participant_name}_{participant_id}_{vibro_label}_ID{difficulty}_trial{trial_count[participant_id][difficulty]}.xlsx"
    filepath = os.path.join(participant_folder, filename)
    wb.save(filepath)

    summary = f"Avg MT: {avg_mt:.2f} ms\nAvg Speed: {avg_speed:.2f} px/ms\nAvg Throughput: {avg_tp:.2f} bit/s\nAvg Error: {avg_err*100:.2f}%"
    if messagebox.askyesno("Experiment Finished", f"{summary}\n\nDo you want to continue with another difficulty?", parent=experiment_window):
        experiment_finished = False
        begin_trial()
    else:
        experiment_window.destroy()

def end_trial():
    save_data_and_finish()

def listen_serial():
    while True:
        if ser and ser.in_waiting:
            line = ser.readline().decode('utf-8').strip()
            if line == '1':
                experiment_window.event_generate('<<SerialClick>>', when='tail')

def begin_trial():
    global difficulty, clicks, data, previous_rects
    try:
        difficulty = int(simpledialog.askstring("Difficulty", "Enter difficulty (1-5):", parent=experiment_window))
        if difficulty < 1 or difficulty > 5:
            raise ValueError
    except:
        messagebox.showerror("Error", "Invalid difficulty.", parent=experiment_window)
        return
    clicks = 0
    data.clear()
    previous_rects.clear()
    trial_count.setdefault(participant_id, {}).setdefault(difficulty, 0)
    trial_count[participant_id][difficulty] += 1
    canvas.pack()
    draw_rectangle()

def start_experiment():
    global participant_name, participant_id, vibro_feedback, participant_folder
    participant_name = simpledialog.askstring("Participant", "Enter Name:", parent=experiment_window)
    participant_id = simpledialog.askstring("Participant", "Enter ID:", parent=experiment_window)
    vibro_feedback = messagebox.askyesno("Vibrotactile", "Enable vibrotactile feedback?", parent=experiment_window)
    participant_folder = f"{participant_name}_{participant_id}"
    if not os.path.exists(participant_folder):
        os.makedirs(participant_folder)
    with open("current_session_path.txt", "w") as f:
        f.write(participant_folder)
    begin_trial()

btn_start = tk.Button(experiment_window, text="Start Experiment", command=start_experiment, font=("Arial", 20))
btn_start.pack(pady=10)
canvas.unbind("<Button-1>")

if ser:
    serial_thread = threading.Thread(target=listen_serial, daemon=True)
    serial_thread.start()

experiment_window.bind('<<SerialClick>>', handle_serial_click)
experiment_window.mainloop()
