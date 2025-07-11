# Python version of Fitts' Law experiment with Teensy serial button input
# GUI implementation using tkinter and canvas for rectangles

import tkinter as tk
from tkinter import simpledialog, messagebox
import time
import openpyxl
from openpyxl import Workbook
import os
import serial
import threading
#import winsound


# Constants
W_VALUES = [90, 80, 70, 60, 50]
D_VALUES = [90, 160, 280, 480, 800]
TOTAL_TRIALS = 20
CANVAS_WIDTH = 1800
CANVAS_HEIGHT = 900
SERIAL_PORT = 'COM7'  # Update with correct Teensy COM port
BAUD_RATE = 9600

# Initialize experiment variables
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
mouse_position = (0, 0)

# Initialize serial connection
try:
    ser = serial.Serial(SERIAL_PORT, BAUD_RATE, timeout=0.1)
except:
    ser = None
    print("Warning: Serial port not available.")

# GUI Setup
root = tk.Tk()
root.title("Fitts' Law Experiment")
canvas = tk.Canvas(root, width=CANVAS_WIDTH, height=CANVAS_HEIGHT, bg="black")
canvas.pack()
canvas.pack_forget()  # Hidden until experiment starts

# Update mouse position on motion
def track_mouse(event):
    global mouse_position
    mouse_position = (event.x, event.y)

canvas.bind("<Motion>", track_mouse)

# Draw rectangle
def draw_rectangle():
    global rect_x, rect_y, rect_width, rect_height, distance, start_time
    canvas.delete("all")

    for rect in previous_rects:
        canvas.create_rectangle(rect[0], rect[1], rect[0]+rect[2], rect[1]+rect[3], fill="lightgrey")

    rect_width = W_VALUES[difficulty - 1]
    distance = D_VALUES[difficulty - 1]
    rect_height = CANVAS_HEIGHT


    center_x = CANVAS_WIDTH / 2
    rect_x = center_x + (target_side * D_VALUES[difficulty - 1]) - (rect_width / 2)
    rect_y = 0

    canvas.create_rectangle(rect_x, rect_y, rect_x + rect_width, rect_y + rect_height, fill="blue")
    start_time = time.time()

# Handle button click event (from serial)
def handle_serial_click(event=None):
    global mouse_position
    x, y = mouse_position
    
    # 🔔 Play beep sound: frequency 1000 Hz, duration 100 ms
    #winsound.Beep(1000, 100)

    mock_event = type('Event', (object,), {'x': x, 'y': y})()
    on_click(mock_event)


# Serial thread function
def listen_serial():
    while True:
        if ser and ser.in_waiting:
            line = ser.readline().decode('utf-8').strip()
            if line == '1':
                root.event_generate('<<SerialClick>>', when='tail')

# Click handler
def on_click(event):
    global clicks, target_side, rect_x, rect_y, rect_width, rect_height
    click_time = time.time()
    mt = (click_time - start_time) * 1000
    error = not (rect_x <= event.x <= rect_x + rect_width)
    speed = distance / mt
    throughput = difficulty / (mt / 1000)

    data.append({
        'MT': mt,
        'speed': speed,
        'error': int(error),
        'clickX': event.x,
        'clickY': event.y,
        'throughput': throughput
    })

    previous_rects.append((rect_x, rect_y, rect_width, rect_height))
    clicks += 1
    target_side *= -1

    if clicks >= TOTAL_TRIALS:
        end_trial()
    else:
        draw_rectangle()

# Start a new trial
def begin_trial():
    global difficulty, rect_width, rect_height, clicks, data, previous_rects
    try:
        difficulty = int(simpledialog.askstring("Difficulty", "Enter difficulty (1-5):"))
        if difficulty < 1 or difficulty > 5:
            raise ValueError
    except:
        messagebox.showerror("Error", "Invalid difficulty.")
        return

    rect_width = W_VALUES[difficulty - 1]
    rect_height = CANVAS_HEIGHT
    clicks = 0
    data.clear()
    previous_rects.clear()

    trial_count.setdefault(participant_id, {}).setdefault(difficulty, 0)
    trial_count[participant_id][difficulty] += 1

    canvas.pack()
    draw_rectangle()

# End the current trial and save results
def end_trial():
    avg_mt = sum(d['MT'] for d in data) / TOTAL_TRIALS
    avg_speed = sum(d['speed'] for d in data) / TOTAL_TRIALS
    avg_tp = sum(d['throughput'] for d in data) / TOTAL_TRIALS
    avg_err = sum(d['error'] for d in data) / TOTAL_TRIALS

    data.append({
        'MT': avg_mt,
        'speed': avg_speed,
        'error': avg_err,
        'clickX': 'Avg',
        'clickY': '',
        'throughput': avg_tp
    })

    wb = Workbook()
    ws = wb.active
    ws.title = "Results"
    headers = ['MT', 'speed', 'error', 'clickX', 'clickY', 'throughput']
    ws.append(headers)
    for row in data:
        ws.append([row.get(h, '') for h in headers])

    vibro_label = "VibroOn" if vibro_feedback else "VibroOff"
    filename = f"{participant_name}_{participant_id}_{vibro_label}_ID{difficulty}_trial{trial_count[participant_id][difficulty]}.xlsx"
    filepath = os.path.join(participant_folder, filename)
    wb.save(filepath)

    messagebox.showinfo("Trial Finished", f"Avg MT: {avg_mt:.2f} ms\nAvg Speed: {avg_speed:.2f} px/ms\nAvg Throughput: {avg_tp:.2f} bit/s\nAvg Error: {avg_err*100:.2f}%")
    canvas.pack_forget()
    ask_next_difficulty()

# Ask for next difficulty or finish
def ask_next_difficulty():
    if messagebox.askyesno("Continue", "Do you want to run another difficulty for this participant?"):
        begin_trial()
    else:
        finish_experiment()

# Start Experiment dialog
def start_experiment():
    global participant_name, participant_id, vibro_feedback, participant_folder
    participant_name = simpledialog.askstring("Participant", "Enter Name:")
    participant_id = simpledialog.askstring("Participant", "Enter ID:")
    vibro_feedback = messagebox.askyesno("Vibrotactile", "Enable vibrotactile feedback?")

    participant_folder = f"{participant_name}_{participant_id}"
    if not os.path.exists(participant_folder):
        os.makedirs(participant_folder)

    begin_trial()

# Finish experiment completely
def finish_experiment():
    root.quit()

# Button setup
btn_start = tk.Button(root, text="Start Experiment", command=start_experiment, font=("Arial", 20))
btn_start.pack(pady=10)

# Remove mouse click binding
tk.Canvas.unbind(canvas, "<Button-1>")

root.bind('<<SerialClick>>', handle_serial_click)

# Start serial listener thread if serial is available
if ser:
    serial_thread = threading.Thread(target=listen_serial, daemon=True)
    serial_thread.start()

root.mainloop()
