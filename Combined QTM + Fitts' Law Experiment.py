# Combined QTM + Fitts' Law Experiment with Forced Second Screen Placement + Pen Tip Logging

import time
import csv
import asyncio
from datetime import datetime
import tkinter as tk
from tkinter import simpledialog, messagebox
import openpyxl
from openpyxl import Workbook
import os
import serial
import threading
import numpy as np
import qtm_rt
from qtm_rt.packet import QRTComponentType
from pythonosc.udp_client import SimpleUDPClient
from screeninfo import get_monitors
import ctypes

# --- QTM configuration ---
QTM_HOST = '127.0.0.1'
OSC_HOST = '139.19.40.134'
UDP_port = 12345
client = SimpleUDPClient(OSC_HOST, UDP_port)
osc_address = '/qtm'
DURATION_SECONDS = 1000
OUTPUT_FILE = 'new.csv'
latest_qtm = None
screen_corners = []

# Fitts' Law Experiment Constants
W_VALUES = [90, 80, 70, 60, 50]
D_VALUES = [90, 160, 280, 480, 800]
TOTAL_TRIALS = 20
SERIAL_PORT = 'COM7'
BAUD_RATE = 9600
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
mm_to_px_x = 1.0
mm_to_px_y = 1.0
pen_tip_xyz = (0.0, 0.0, 0.0)
last_pen_local = (0.0, 0.0)
current_rect = (0, 0, 0, 0)

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

# QTM Geometry Functions
def rect_point_to_local_xy(corners, point):
    corners = np.array(corners)
    point = np.array(point)
    center = np.mean(corners, axis=0)
    u = corners[1] - corners[0]
    v = corners[3] - corners[0]
    u_norm = u / np.linalg.norm(u)
    v_norm = v / np.linalg.norm(v)
    vec = point - center
    x_local = np.dot(vec, u_norm)
    y_local = np.dot(vec, v_norm)
    return x_local, y_local

def distance_point_to_plane(point, plane_point, plane_normal):
    point = np.array(point)
    plane_point = np.array(plane_point)
    plane_normal = np.array(plane_normal)
    plane_normal = plane_normal / np.linalg.norm(plane_normal)
    return np.dot(point - plane_point, plane_normal)

def handle_qtm_data(packet):
    global latest_qtm, screen_corners, mm_to_px_x, mm_to_px_y, pen_tip_xyz, last_pen_local
    try:
        frame = packet.framenumber
        latest_qtm = []
        header, markers = packet.get_3d_markers()
        labels = ["Pen - 1", "Pen Tip", "Pen - 3", "Pen - 4", "Screen - 1", "Screen - 2", "Screen - 3", "Screen - 4"]

        for i, marker in enumerate(markers):
            latest_qtm.append([marker.x, marker.y, marker.z])

        if len(latest_qtm) >= 8:
            screen_corners = latest_qtm[4:8]
            pen_tip_xyz = tuple(latest_qtm[1])

            v1 = np.array(screen_corners[1]) - np.array(screen_corners[0])
            v2 = np.array(screen_corners[3]) - np.array(screen_corners[0])
            normal = np.cross(v1, v2)
            normal /= np.linalg.norm(normal)

            screen_width_mm = np.linalg.norm(v1)
            screen_height_mm = np.linalg.norm(v2)
            mm_to_px_x = CANVAS_WIDTH / screen_width_mm
            mm_to_px_y = CANVAS_HEIGHT / screen_height_mm

            # Project to local coordinates
            x_local, y_local = rect_point_to_local_xy(screen_corners, pen_tip_xyz)
            last_pen_local = (x_local * mm_to_px_x + CANVAS_WIDTH/2, y_local * mm_to_px_y + CANVAS_HEIGHT/2)

    except Exception as e:
        print(f"bug {e}")
        latest_qtm = ['NaN', 'NaN', 'NaN']

async def start_qtm_listener():
    conn = await qtm_rt.connect(QTM_HOST)
    await conn.stream_frames(components=['3d'], on_packet=handle_qtm_data)

async def wait_for_valid_qtm():
    while True:
        await asyncio.sleep(0.01)
        if screen_corners and len(screen_corners) == 4:
            break

# Replace handle_serial_click to capture and log tip location

def handle_serial_click(event=None):
    global clicks, target_side, current_rect
    click_time = time.time()
    mt = (click_time - start_time) * 1000
    x_pen, y_pen = last_pen_local

    rect_x, rect_y, rect_w, rect_h = current_rect
    inside = rect_x <= x_pen <= rect_x + rect_w and rect_y <= y_pen <= rect_y + rect_h
    error = not inside
    speed = distance / mt
    throughput = difficulty / (mt / 1000)

    data.append({
        'MT': mt,
        'speed': speed,
        'error': int(error),
        'clickX': int(inside),
        'clickY': 0,
        'throughput': throughput,
        'PenX': pen_tip_xyz[0],
        'PenY': pen_tip_xyz[1],
        'PenZ': pen_tip_xyz[2]
    })

    previous_rects.append(current_rect)
    clicks += 1
    target_side *= -1

    if clicks >= TOTAL_TRIALS:
        end_trial()
    else:
        draw_rectangle()

experiment_window.bind('<<SerialClick>>', handle_serial_click)

def listen_serial():
    while True:
        if ser and ser.in_waiting:
            line = ser.readline().decode('utf-8').strip()
            if line == '1':
                experiment_window.event_generate('<<SerialClick>>', when='tail')

# -- Fitts GUI logic --
def draw_rectangle():
    global rect_x, rect_y, rect_width, rect_height, distance, start_time, current_rect
    canvas.delete("all")
    for rect in previous_rects:
        canvas.create_rectangle(rect[0], rect[1], rect[0]+rect[2], rect[1]+rect[3], fill="lightgrey")
    rect_width = W_VALUES[difficulty - 1]
    distance = D_VALUES[difficulty - 1]
    rect_height = CANVAS_HEIGHT
    center_x = CANVAS_WIDTH / 2
    rect_x = center_x + (target_side * distance) - (rect_width / 2)
    rect_y = 0
    canvas.create_rectangle(rect_x, rect_y, rect_x + rect_width, rect_y + rect_height, fill="blue")
    current_rect = (rect_x, rect_y, rect_width, rect_height)
    start_time = time.time()

def begin_trial():
    global difficulty, rect_width, rect_height, clicks, data, previous_rects
    try:
        difficulty = int(simpledialog.askstring("Difficulty", "Enter difficulty (1-5):", parent=experiment_window))
        if difficulty < 1 or difficulty > 5:
            raise ValueError
    except:
        messagebox.showerror("Error", "Invalid difficulty.", parent=experiment_window)
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
        'throughput': avg_tp,
        'PenX': '',
        'PenY': '',
        'PenZ': ''
    })
    wb = Workbook()
    ws = wb.active
    ws.title = "Results"
    headers = ['MT', 'speed', 'error', 'clickX', 'clickY', 'throughput', 'PenX', 'PenY', 'PenZ']
    ws.append(headers)
    for row in data:
        ws.append([row.get(h, '') for h in headers])
    vibro_label = "VibroOn" if vibro_feedback else "VibroOff"
    filename = f"{participant_name}_{participant_id}_{vibro_label}_ID{difficulty}_trial{trial_count[participant_id][difficulty]}.xlsx"
    filepath = os.path.join(participant_folder, filename)
    wb.save(filepath)
    messagebox.showinfo("Trial Finished", f"Avg MT: {avg_mt:.2f} ms\nAvg Speed: {avg_speed:.2f} px/ms\nAvg Throughput: {avg_tp:.2f} bit/s\nAvg Error: {avg_err*100:.2f}%", parent=experiment_window)
    canvas.pack_forget()
    ask_next_difficulty()

def ask_next_difficulty():
    if messagebox.askyesno("Continue", "Do you want to run another difficulty for this participant?", parent=experiment_window):
        begin_trial()
    else:
        finish_experiment()

def start_experiment():
    global participant_name, participant_id, vibro_feedback, participant_folder
    participant_name = simpledialog.askstring("Participant", "Enter Name:", parent=experiment_window)
    participant_id = simpledialog.askstring("Participant", "Enter ID:", parent=experiment_window)
    vibro_feedback = messagebox.askyesno("Vibrotactile", "Enable vibrotactile feedback?", parent=experiment_window)
    participant_folder = f"{participant_name}_{participant_id}"
    if not os.path.exists(participant_folder):
        os.makedirs(participant_folder)
    begin_trial()

def finish_experiment():
    experiment_window.quit()

btn_start = tk.Button(experiment_window, text="Start Experiment", command=start_experiment, font=("Arial", 20))
btn_start.pack(pady=10)
canvas.unbind("<Button-1>")

if ser:
    serial_thread = threading.Thread(target=listen_serial, daemon=True)
    serial_thread.start()

def run_tk():
    experiment_window.mainloop()

async def main():
    asyncio.ensure_future(start_qtm_listener())
    await wait_for_valid_qtm()
    run_tk()

if __name__ == '__main__':
    time.sleep(2)
    asyncio.run(main())
