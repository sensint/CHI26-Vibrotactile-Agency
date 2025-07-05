import time
import csv
import asyncio
from datetime import datetime

import qtm_rt
from qtm_rt.packet import QRTComponentType
from pythonosc.udp_client import SimpleUDPClient
import numpy as np

# --- QTM configuration ---
QTM_HOST = '127.0.0.1'  # QTM server address
OSC_HOST = '139.19.40.134'
UDP_port = 12345
client = SimpleUDPClient(OSC_HOST, UDP_port)
osc_address = '/qtm'

# --- Duration for logging (in seconds) ---
DURATION_SECONDS = 1000

# --- Output file ---
OUTPUT_FILE = 'new.csv'

latest_qtm = None

def rect_point_to_local_xy(corners, point):
    """
    Convert 3D point on rectangle to local 2D coordinates in rectangle frame.
    """
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
    """
    Computes signed distance from `point` to a plane.
    """
    point = np.array(point)
    plane_point = np.array(plane_point)
    plane_normal = np.array(plane_normal)
    plane_normal = plane_normal / np.linalg.norm(plane_normal)
    return np.dot(point - plane_point, plane_normal)

def handle_qtm_data(packet):
    global latest_qtm

    try:
        frame = packet.framenumber
        latest_qtm = []

        print(f"\nFramenumber: {frame}")
        header, markers = packet.get_3d_markers()

        labels = [
            "Pen - 1", "Pen Tip", "Pen - 3", "Pen - 4",
            "Screen - 1", "Screen - 2", "Screen - 3", "Screen - 4"
        ]

        for i, marker in enumerate(markers):
            label = labels[i] if i < len(labels) else f"Marker {i+1}"
            print(f"{label}: X={marker.x:.2f}, Y={marker.y:.2f}, Z={marker.z:.2f}")
            latest_qtm.append([marker.x, marker.y, marker.z])

        if len(latest_qtm) >= 8:
            screen_corners = latest_qtm[4:8]
            pen_tip = latest_qtm[1]

            # 1. Compute plane normal from screen corners
            v1 = np.array(screen_corners[1]) - np.array(screen_corners[0])
            v2 = np.array(screen_corners[3]) - np.array(screen_corners[0])
            normal = np.cross(v1, v2)
            normal /= np.linalg.norm(normal)

            # 2. Compute perpendicular distance to screen plane
            plane_point = screen_corners[0]
            dist = abs(distance_point_to_plane(pen_tip, plane_point, normal))

            print(f"Pen distance to screen: {dist:.2f} mm")

            # 3. Check if pen is "touching"
            if dist < 4.0:  # threshold in mm
                x_local, y_local = rect_point_to_local_xy(screen_corners, pen_tip)
                print(f"✅ Pen Tip TOUCHING screen at local: x={x_local:.2f}, y={y_local:.2f}")
            else:
                print("🛑 Pen is not touching the screen.")

    except Exception as e:
        print(f"bug {e}")
        latest_qtm = ['NaN', 'NaN', 'NaN', 'NaN']

async def start_qtm_listener():
    conn = await qtm_rt.connect(QTM_HOST)
    await conn.stream_frames(components=['3d', '6d'], on_packet=handle_qtm_data)

async def wait_for_valid_qtm():
    while True:
        await asyncio.sleep(0.01)
        if latest_qtm and isinstance(latest_qtm[0], float):
            break

async def main():
    global latest_qtm
    asyncio.ensure_future(start_qtm_listener())
    await wait_for_valid_qtm()

    with open(OUTPUT_FILE, mode='w', newline='') as file:
        writer = csv.writer(file)
        start_dt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        writer.writerow([f"Start Time: {start_dt}"])
        writer.writerow([])
        writer.writerow(['time_python(ms)', 'qtm_x', 'qtm_y', 'qtm_z', 'qtm_rot'])

        start_time = time.time()

        for elapsed_ms in range(DURATION_SECONDS * 1000):
            now = time.time()
            to_wait = (start_time + (elapsed_ms + 1) / 1000) - now
            if to_wait > 0:
                await asyncio.sleep(to_wait)

            # Could add logging of pen tip XY here if desired

        end_dt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        writer.writerow([f"End Time: {end_dt}"])

if __name__ == '__main__':
    time.sleep(2)
    asyncio.run(main())
