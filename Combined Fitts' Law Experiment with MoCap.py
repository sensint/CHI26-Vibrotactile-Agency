# Combined QTM + Fitts' Law Experiment

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
CANVAS_WIDTH = 1800
CANVAS_HEIGHT = 900
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

try:
    ser = serial.Serial(SERIAL_PORT, BAUD_RATE, timeout=0.1)
except:
    ser = None
    print("Warning: Serial port not available.")

root = tk.Tk()
root.title("Fitts' Law Experiment")
canvas = tk.Canvas(root, width=CANVAS_WIDTH, height=CANVAS_HEIGHT, bg="white")
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
    global latest_qtm, screen_corners
    try:
        frame = packet.framenumber
        latest_qtm = []
        header, markers = packet.get_3d_markers()

        labels = ["Pen - 1", "Pen Tip", "Pen - 3", "Pen - 4", "Screen - 1", "Screen - 2", "Screen - 3", "Screen - 4"]

        for i, marker in enumerate(markers):
            label = labels[i] if i < len(labels) else f"Marker {i+1}"
            print(f"{label}: X={marker.x:.2f}, Y={marker.y:.2f}, Z={marker.z:.2f}")
            latest_qtm.append([marker.x, marker.y, marker.z])

        if len(latest_qtm) >= 8:
            screen_corners = latest_qtm[4:8]
            pen_tip = latest_qtm[1]

            v1 = np.array(screen_corners[1]) - np.array(screen_corners[0])
            v2 = np.array(screen_corners[3]) - np.array(screen_corners[0])
            normal = np.cross(v1, v2)
            normal /= np.linalg.norm(normal)

            plane_point = screen_corners[0]
            dist = abs(distance_point_to_plane(pen_tip, plane_point, normal))

            print(f"Pen distance to screen: {dist:.2f} mm")

            if dist < 4.0:
                x_local, y_local = rect_point_to_local_xy(screen_corners, pen_tip)
                print(f"✅ Pen Tip TOUCHING screen at local: x={x_local:.2f}, y={y_local:.2f}")
            else:
                print("🛑 Pen is not touching the screen.")

    except Exception as e:
        print(f"bug {e}")
        latest_qtm = ['NaN', 'NaN', 'NaN', 'NaN']

async def start_qtm_listener():
    conn = await qtm_rt.connect(QTM_HOST)
    await conn.stream_frames(components=['3d'], on_packet=handle_qtm_data)

async def wait_for_valid_qtm():
    while True:
        await asyncio.sleep(0.01)
        if screen_corners and len(screen_corners) == 4:
            break

# FITTS GUI Functions
def track_mouse(event):
    global mouse_position
    mouse_position = (event.x, event.y)
canvas.bind("<Motion>", track_mouse)

# Modify draw_rectangle to calculate screen-relative positions using QTM corners

def draw_rectangle():
    global rect_x, rect_y, rect_width, rect_height, distance, start_time
    canvas.delete("all")

    for rect in previous_rects:
        canvas.create_rectangle(rect[0], rect[1], rect[0]+rect[2], rect[1]+rect[3], fill="lightgrey")

    rect_width = W_VALUES[difficulty - 1]
    distance = D_VALUES[difficulty - 1]
    rect_height = CANVAS_HEIGHT

    if len(screen_corners) == 4:
        # QTM-based 3D positioning of target rectangle in screen plane
        corners = np.array(screen_corners)
        u = corners[1] - corners[0]  # horizontal axis
        u /= np.linalg.norm(u)
        center = np.mean(corners, axis=0)

        # Compute the 3D center of rectangle
        rect_center_3d = center + target_side * u * distance
        x1_3d = rect_center_3d - (u * rect_width / 2)
        x2_3d = rect_center_3d + (u * rect_width / 2)

        print(f"🎯 3D Rectangle span: From {x1_3d} to {x2_3d}, width: {rect_width}px, offset: {distance}px")

    # Fallback 2D canvas display
    center_x = CANVAS_WIDTH / 2
    rect_x = center_x + (target_side * distance) - (rect_width / 2)
    rect_y = 0
    canvas.create_rectangle(rect_x, rect_y, rect_x + rect_width, rect_y + rect_height, fill="blue")
    start_time = time.time()

def handle_serial_click(event=None):
    global mouse_position
    x, y = mouse_position
    mock_event = type('Event', (object,), {'x': x, 'y': y})()
    on_click(mock_event)

def listen_serial():
    while True:
        if ser and ser.in_waiting:
            line = ser.readline().decode('utf-8').strip()
            if line == '1':
                root.event_generate('<<SerialClick>>', when='tail')

def on_click(event):
    global clicks, target_side, rect_x, rect_y, rect_width, rect_height
    click_time = time.time()
    mt = (click_time - start_time) * 1000
    error = not (rect_x <= event.x <= rect_x + rect_width)
    speed = distance / mt
    throughput = difficulty / (mt / 1000)

    data.append({'MT': mt, 'speed': speed, 'error': int(error), 'clickX': event.x, 'clickY': event.y, 'throughput': throughput})
    previous_rects.append((rect_x, rect_y, rect_width, rect_height))
    clicks += 1
    target_side *= -1

    if clicks >= TOTAL_TRIALS:
        end_trial()
    else:
        draw_rectangle()

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

def end_trial():
    avg_mt = sum(d['MT'] for d in data) / TOTAL_TRIALS
    avg_speed = sum(d['speed'] for d in data) / TOTAL_TRIALS
    avg_tp = sum(d['throughput'] for d in data) / TOTAL_TRIALS
    avg_err = sum(d['error'] for d in data) / TOTAL_TRIALS

    data.append({'MT': avg_mt, 'speed': avg_speed, 'error': avg_err, 'clickX': 'Avg', 'clickY': '', 'throughput': avg_tp})
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

def ask_next_difficulty():
    if messagebox.askyesno("Continue", "Do you want to run another difficulty for this participant?"):
        begin_trial()
    else:
        finish_experiment()

def start_experiment():
    global participant_name, participant_id, vibro_feedback, participant_folder
    participant_name = simpledialog.askstring("Participant", "Enter Name:")
    participant_id = simpledialog.askstring("Participant", "Enter ID:")
    vibro_feedback = messagebox.askyesno("Vibrotactile", "Enable vibrotactile feedback?")

    participant_folder = f"{participant_name}_{participant_id}"
    if not os.path.exists(participant_folder):
        os.makedirs(participant_folder)

    begin_trial()

def finish_experiment():
    root.quit()

btn_start = tk.Button(root, text="Start Experiment", command=start_experiment, font=("Arial", 20))
btn_start.pack(pady=10)
canvas.unbind("<Button-1>")
root.bind('<<SerialClick>>', handle_serial_click)
if ser:
    serial_thread = threading.Thread(target=listen_serial, daemon=True)
    serial_thread.start()

async def main():
    asyncio.ensure_future(start_qtm_listener())
    await wait_for_valid_qtm()
    root.mainloop()

if __name__ == '__main__':
    time.sleep(2)
    asyncio.run(main())
