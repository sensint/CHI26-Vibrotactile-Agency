import time
import asyncio
import csv
import subprocess
import numpy as np
from pathlib import Path
import qtm_rt
from qtm_rt.packet import QRTComponentType
import sys
from datetime import datetime
from pythonosc.udp_client import SimpleUDPClient
import zmq
import json
import nidaqmx
from nidaqmx.constants import AcquisitionType

# Configuration - QTM and Logging
QTM_HOST = '139.19.40.134'
OSC_HOST = '139.19.40.35'
UDP_port = 12345
ZMQ_PORT = 5555
ZMQ_CONFIG_PORT = 5556  # Port for sending config to subscriber
ZMQ_TOPIC = "qtm_data"

# Configuration - Motion-triggered vibration
FSR_MIN = 0
FSR_MAX = 344
NUM_BINS = 100
FREQUENCY = 50
AMPLITUDE = 1  # V
FS_OUTPUT = 5000  # Output sample rate (Hz)
DEVICE_AO = "Dev1/ao0"  # Analog output
HYSTERESIS_PERCENT = 0.3  # Hysteresis as fraction of bin width (0.3 = 30%)

# Marker indices (0-based) — set these to match your QTM marker setup
MARKER_TOP_RIGHT = 0     # Index for top right screen corner marker
MARKER_BOTTOM_RIGHT = 1  # Index for bottom right screen corner marker
MARKER_BOTTOM_LEFT = 2   # Index for bottom left screen corner marker
MARKER_TOP_LEFT = 3      # Index for top left screen corner marker
MARKER_PEN_TIP = 8       # Index for pen tip marker
# Minimum number of markers expected from QTM
MIN_MARKERS = max(MARKER_TOP_RIGHT, MARKER_BOTTOM_RIGHT, MARKER_BOTTOM_LEFT, MARKER_TOP_LEFT, MARKER_PEN_TIP) + 1

# Vibration condition: "motion-coupled" | "continuous" | "no-vibration"
CONDITION = "continuous"
participant_name = "test"
attempts = 2
ID = 2                    # Set to 2 or 4 (Index of Difficulty)
delaytime = 250

# ID-to-parameters lookup: W (px), D (px)
ID_PARAMS = {
    2: {"W": 80, "D": 240},    # log2(240/80 + 1) = 2 bits
    4: {"W": 60, "D": 900},    # log2(900/60 + 1) = 4 bits
}

if ID not in ID_PARAMS:
    raise ValueError(f"Unsupported ID={ID}. Supported values: {list(ID_PARAMS.keys())}")

# Generate single-cycle sine wave for output
samples_per_period = int(FS_OUTPUT / FREQUENCY)  # 50 samples for 1 cycle at 5kHz
t = np.linspace(0, 1/FREQUENCY, samples_per_period, endpoint=False)
single_cycle_wave = AMPLITUDE * np.sin(2 * np.pi * FREQUENCY * t)
# Append a zero sample at the end to reset output to 0V
single_cycle_wave = np.append(single_cycle_wave, 0.0)

# Global variables
task_ao = None
last_bin = -1
last_trigger_x = None  # Hysteresis: track x_local at last bin change
start_time = None
streaming_enabled = True
continuous_playing = False  # For continuous mode: is DAQ currently outputting?

# Trigger detection (replicated from subscriber)
W_VALUES = [ID_PARAMS[ID]["W"]]       # Target width in pixels (derived from ID)
D_VALUES = [ID_PARAMS[ID]["D"]]       # Distance between targets in pixels (derived from ID)
TOTAL_TRIALS = 10
SCREEN_WIDTH_MM = 346.0   # Physical screen width in mm (default)
CANVAS_WIDTH = 1920       # Screen resolution width in pixels (adjust to your monitor)
target_side = 1           # 1 = right, -1 = left
trigger_count = 0
previous_inside = False
rect_x_mm = None
rect_x_end_mm = None

client = SimpleUDPClient(OSC_HOST, UDP_port)
osc_address = '/qtm'

session_file = Path(__file__).parent / "current_session_path.txt"

log_dir = None  # Set in main()
log_rows = []
clicked_frames = set()
latest_frame = None

# Set vibration mode from CONDITION
vibration_mode = CONDITION
print(f"🔧 Vibration mode: {vibration_mode}")


def calculate_target_bounds():
    """Calculate target bounds in mm based on current target_side, replicating subscriber logic."""
    global rect_x_mm, rect_x_end_mm
    rect_width = W_VALUES[0]
    distance = D_VALUES[0]
    center_x = CANVAS_WIDTH / 2
    rect_x = center_x + (target_side * (distance / 2)) - (rect_width / 2)
    px_to_mm = SCREEN_WIDTH_MM / CANVAS_WIDTH
    rect_x_mm = rect_x * px_to_mm
    rect_x_end_mm = (rect_x + rect_width) * px_to_mm
    print(f"Target bounds (mm): {rect_x_mm:.2f} to {rect_x_end_mm:.2f} | Side: {'right' if target_side == 1 else 'left'}")


# Initialize first target bounds
calculate_target_bounds()

# ZeroMQ setup
zmq_context = zmq.Context()
zmq_socket = zmq_context.socket(zmq.PUB)
zmq_socket.bind(f"tcp://*:{ZMQ_PORT}")
print(f"✅ ZeroMQ publisher started on port {ZMQ_PORT}")


def initialize_daq():
    """Initialize NI-DAQ analog output based on vibration_mode"""
    global task_ao, start_time
    
    if vibration_mode == "no-vibration":
        print("DAQ: no-vibration mode — DAQ not initialized")
        start_time = time.time()
        return True
    
    try:
        # Close any existing task
        if task_ao is not None:
            try:
                task_ao.close()
            except:
                pass
        
        task_ao = nidaqmx.Task()
        task_ao.ao_channels.add_ao_voltage_chan(
            DEVICE_AO,
            min_val=-10.0,
            max_val=10.0
        )

        if vibration_mode == "continuous":
            # Continuous mode: repeating sine wave, started/stopped on touch
            task_ao.timing.cfg_samp_clk_timing(
                rate=FS_OUTPUT,
                sample_mode=AcquisitionType.CONTINUOUS,
                samps_per_chan=samples_per_period
            )
            task_ao.write(single_cycle_wave[:-1], auto_start=False)  # exclude trailing zero for seamless looping
            print(f"DAQ configured (continuous mode):")
            print(f"  Output: {FREQUENCY}Hz sine wave, ±{AMPLITUDE}V (loops while touching)")
        else:
            # Motion-coupled mode: single-cycle burst per bin change
            task_ao.timing.cfg_samp_clk_timing(
                rate=FS_OUTPUT,
                sample_mode=AcquisitionType.FINITE,
                samps_per_chan=len(single_cycle_wave)
            )
            task_ao.write(single_cycle_wave, auto_start=False)
            print(f"DAQ configured (motion-coupled mode):")
            print(f"  Output: {FREQUENCY}Hz sine wave, ±{AMPLITUDE}V (burst per bin change)")
        
        print(f"  Sample rate: {FS_OUTPUT}Hz ({samples_per_period} samples/cycle)")
        start_time = time.time()
        return True
        
    except Exception as e:
        print(f"Error initializing DAQ: {e}")
        return False


def cleanup_daq():
    """Cleanup DAQ resources"""
    global task_ao, continuous_playing
    
    if vibration_mode == "no-vibration":
        return
    
    try:
        if task_ao is not None:
            if continuous_playing:
                task_ao.stop()
                continuous_playing = False
            task_ao.close()
            print("DAQ task closed")
    except Exception as e:
        print(f"DAQ cleanup warning: {e}")


def start_continuous():
    """Start continuous sine wave output (for continuous mode)"""
    global continuous_playing
    if not continuous_playing and task_ao is not None:
        try:
            task_ao.start()
            continuous_playing = True
        except Exception as e:
            print(f"Continuous start error: {e}")


def stop_continuous():
    """Stop continuous sine wave output (for continuous mode)"""
    global continuous_playing
    if continuous_playing and task_ao is not None:
        try:
            task_ao.stop()
            continuous_playing = False
        except Exception as e:
            print(f"Continuous stop error: {e}")


def trigger_burst():
    """Trigger single cycle output (motion-coupled mode only)"""
    global task_ao
    
    if vibration_mode != "motion-coupled" or task_ao is None:
        return
    
    try:
        task_ao.start()
        task_ao.wait_until_done(timeout=1.0)
        task_ao.stop()
    except Exception as e:
        print(f"Trigger error: {e}")


def map_value(value, in_min, in_max, out_min, out_max):
    """Arduino-style map function"""
    return int((value - in_min) * (out_max - out_min) / (in_max - in_min) + out_min)


def calculate_screen_plane_normal(screen_corners):
    """Calculate the normal vector of the screen plane"""
    p0 = np.array(screen_corners[0])
    p1 = np.array(screen_corners[1])
    p3 = np.array(screen_corners[3])
    
    # Two vectors in the plane
    v1 = p1 - p0
    v2 = p3 - p0
    
    # Normal is cross product
    normal = np.cross(v1, v2)
    return normal / np.linalg.norm(normal)


def project_point_to_plane(point, plane_point, plane_normal):
    """Project a point onto a plane defined by a point and normal vector"""
    point = np.array(point)
    plane_point = np.array(plane_point)
    plane_normal = np.array(plane_normal)
    
    # Vector from plane point to the point
    v = point - plane_point
    
    # Distance from point to plane
    dist = np.dot(v, plane_normal)
    
    # Project point onto plane
    projected = point - dist * plane_normal
    
    return projected


def get_distance_from_reference_line(pen_tip, screen_corners):
    """Calculate perpendicular distance from reference line (Point 2-3)
    
    Args:
        pen_tip: [x, y, z] coordinates of pen tip
        screen_corners: list of [x, y, z] coordinates for 4 screen corners
    
    Returns:
        distance: perpendicular distance from reference line in mm
        is_valid: True if projection is within screen bounds
    """
    # Get corner positions
    p0 = np.array(screen_corners[0])
    p1 = np.array(screen_corners[1])
    p2 = np.array(screen_corners[2])
    p3 = np.array(screen_corners[3])
    
    # Calculate screen plane normal
    plane_normal = calculate_screen_plane_normal(screen_corners)
    
    # Project pen tip onto screen plane
    projected_point = project_point_to_plane(np.array(pen_tip), p2, plane_normal)
    
    # Reference line is now from Point 2 to Point 3 (left edge)
    reference_line_vector = p3 - p2  # left edge: bottom left to top left
    reference_line_length = np.linalg.norm(reference_line_vector)
    reference_line_unit = reference_line_vector / reference_line_length

    # Height direction (perpendicular to reference line, in screen plane)
    # This points from left edge to right edge
    height_vector = p1 - p2  # bottom left to bottom right (horizontal)
    # Make it exactly perpendicular by removing component along reference line
    height_vector = height_vector - np.dot(height_vector, reference_line_unit) * reference_line_unit
    height_direction = height_vector / np.linalg.norm(height_vector)
    screen_height = np.linalg.norm(height_vector)

    # Vector from reference line start (Point 2) to projected point
    to_projected = projected_point - p2

    # Distance along reference line (for width validation)
    distance_along_reference = np.dot(to_projected, reference_line_unit)

    # Distance perpendicular to reference line (this is what we want)
    distance_from_reference = np.dot(to_projected, height_direction)

    # Check if projection is within screen bounds
    is_valid = (0 <= distance_from_reference <= screen_height and 
                0 <= distance_along_reference <= reference_line_length)
    
    return distance_from_reference, is_valid


def rect_point_to_local_xy(corners, point):
    """
    Calculate the horizontal distance (x_local) from the left edge (bottom left to bottom right) of the screen rectangle.
    corners: list of 4 points [top right, bottom right, bottom left, top left]
    point: the [x, y, z] coordinates to project (e.g., pen tip)
    Returns:
        x_local: horizontal distance from the left edge (bottom left), in mm
        y_local: always 0 (not used)
    """
    corners = np.array(corners)
    point = np.array(point)
    left_edge_start = corners[2]  # p2 (bottom left)
    right_edge_end = corners[1]   # p1 (bottom right)
    horizontal_vec = right_edge_end - left_edge_start
    horizontal_unit = horizontal_vec / np.linalg.norm(horizontal_vec)
    pen_vec = point - left_edge_start
    x_local = np.dot(pen_vec, horizontal_unit)
    return x_local, 0  # y_local not used


def distance_point_to_plane(point, plane_point, plane_normal):
    point = np.array(point)
    plane_point = np.array(plane_point)
    plane_normal = np.array(plane_normal)
    plane_normal = plane_normal / np.linalg.norm(plane_normal)
    return np.dot(point - plane_point, plane_normal)


def handle_qtm_data(packet):
    global latest_frame, streaming_enabled, last_bin, start_time, last_trigger_x, previous_inside, trigger_count, target_side

    try:
        # Stop streaming after TOTAL_TRIALS triggers
        if trigger_count >= TOTAL_TRIALS:
            return

        frame = packet.framenumber
        latest_frame = frame

        header, markers = packet.get_3d_markers()
        marker_xyz = [[m.x, m.y, m.z] for m in markers]
        
        if len(marker_xyz) >= MIN_MARKERS:
            screen_corners = [
                marker_xyz[MARKER_TOP_RIGHT],
                marker_xyz[MARKER_BOTTOM_RIGHT],
                marker_xyz[MARKER_BOTTOM_LEFT],
                marker_xyz[MARKER_TOP_LEFT],
            ]
            pen_tip = marker_xyz[MARKER_PEN_TIP]

            # Check for invalid markers (NaN or all zeros)
            all_markers_valid = True
            for i in [MARKER_TOP_RIGHT, MARKER_BOTTOM_RIGHT, MARKER_BOTTOM_LEFT, MARKER_TOP_LEFT, MARKER_PEN_TIP]:
                m = marker_xyz[i]
                if (m[0] == 0 and m[1] == 0 and m[2] == 0) or \
                   np.isnan(m[0]) or np.isnan(m[1]) or np.isnan(m[2]):
                    all_markers_valid = False
                    break
            
            if not all_markers_valid:
                return

            # Calculate plane normal and distance
            v1 = np.array(screen_corners[1]) - np.array(screen_corners[0])
            v2 = np.array(screen_corners[3]) - np.array(screen_corners[0])
            normal = np.cross(v1, v2)
            normal /= np.linalg.norm(normal)

            dist = abs(distance_point_to_plane(pen_tip, screen_corners[0], normal))
            
            # Calculate distance from reference line for bin triggering
            distance_from_ref, is_valid_position = get_distance_from_reference_line(
                pen_tip,
                screen_corners
            )
            
            status = "not_touching"
            x_local = y_local = 'NaN'

            if dist < 8.0:
                x_local, y_local = rect_point_to_local_xy(screen_corners, pen_tip)
                if streaming_enabled:
                    client.send_message(osc_address, [round(x_local, 1), round(y_local, 1)])

                width = np.linalg.norm(np.array(screen_corners[1]) - np.array(screen_corners[0]))
                height = np.linalg.norm(np.array(screen_corners[3]) - np.array(screen_corners[0]))
                if -width/2 <= x_local <= width/2 and -height/2 <= y_local <= height/2:
                    status = "touching"
                else:
                    status = "outside"

            # VIBRATION OUTPUT LOGIC (depends on vibration_mode)
            if vibration_mode == "motion-coupled":
                # BIN-BASED TRIGGERING LOGIC with hysteresis
                # Only trigger if pen is close enough to the plane
                BIN_WIDTH = (FSR_MAX - FSR_MIN) / NUM_BINS  # Width of each bin in mm
                HYSTERESIS = BIN_WIDTH * HYSTERESIS_PERCENT
                if is_valid_position and dist < 8.0:
                    # Map x_local to bins
                    fsr_value = int(x_local)
                    current_bin = map_value(fsr_value, FSR_MIN, FSR_MAX, 0, NUM_BINS)
                    current_bin = max(0, min(current_bin, NUM_BINS - 1))
                    if current_bin != last_bin:
                        # Only trigger if x_local has moved enough past the boundary (hysteresis)
                        if last_trigger_x is not None and abs(x_local - last_trigger_x) < HYSTERESIS:
                            pass  # Too close to last trigger point, ignore
                        elif last_bin == -1:
                            last_bin = current_bin
                            last_trigger_x = x_local
                        else:
                            trigger_time = time.time() - start_time
                            trigger_burst()
                            last_bin = current_bin
                            last_trigger_x = x_local
                else:
                    if last_bin != -1:
                        last_bin = -1
                        last_trigger_x = None

            elif vibration_mode == "continuous":
                # CONTINUOUS MODE: output sine wave while pen is touching the screen
                if is_valid_position and dist < 8.0:
                    start_continuous()
                else:
                    stop_continuous()

            # no-vibration: do nothing

            # Send via ZeroMQ
            data = {
                "frame": frame,
                "x_local": round(x_local, 2) if isinstance(x_local, (int, float)) else None,
                "y_local": round(y_local, 2) if isinstance(y_local, (int, float)) else None,
                "distance": round(dist, 2),
                "distance_from_reference": round(distance_from_ref, 2) if is_valid_position else None,
                "current_bin": int(current_bin) if is_valid_position and last_bin != -1 else None,
                "status": status,
                "inside_bounds": int(status == "touching"),
                "is_valid_position": int(is_valid_position),
                "pen_tip": [round(pen_tip[0], 2), round(pen_tip[1], 2), round(pen_tip[2], 2)],
                "screen_corners": [
                    [round(screen_corners[0][0], 2), round(screen_corners[0][1], 2), round(screen_corners[0][2], 2)],
                    [round(screen_corners[1][0], 2), round(screen_corners[1][1], 2), round(screen_corners[1][2], 2)],
                    [round(screen_corners[2][0], 2), round(screen_corners[2][1], 2), round(screen_corners[2][2], 2)],
                    [round(screen_corners[3][0], 2), round(screen_corners[3][1], 2), round(screen_corners[3][2], 2)]
                ]
            }
            message = f"{ZMQ_TOPIC} {json.dumps(data)}"
            zmq_socket.send_string(message)

            # TRIGGER DETECTION (replicated from subscriber)
            # Check if x_local is inside current target bounds
            if x_local is not None and isinstance(x_local, (int, float)) and rect_x_mm is not None and rect_x_end_mm is not None:
                inside_bounds = (rect_x_mm <= x_local <= rect_x_end_mm)
            else:
                inside_bounds = False

            # State Machine: Trigger only on transition from 0 → 1 (outside → inside)
            if inside_bounds and not previous_inside:
                trigger_count += 1
                clicked_frames.add(frame)
                print(f"🔘 TRIGGER {trigger_count}/{TOTAL_TRIALS} | x_local: {x_local:.2f}mm | Target: {rect_x_mm:.2f}-{rect_x_end_mm:.2f}mm")
                # Flip target side for next trial
                target_side *= -1
                calculate_target_bounds()

            previous_inside = inside_bounds

            log_rows.append([
                frame, pen_tip[0], pen_tip[1], pen_tip[2],
                dist, x_local
            ])
            
    except Exception as e:
        print(f"❌ QTM error: {e}")


async def main():
    global log_rows, clicked_frames, start_time, log_dir

    # Initialize DAQ
    if not initialize_daq():
        print("\nFailed to initialize DAQ. Exiting.")
        sys.exit(1)

    # Create results folder and session file
    import os
    base_dir = str(Path(__file__).parent / "Results")
    results_folder = os.path.join(base_dir, participant_name)
    os.makedirs(results_folder, exist_ok=True)
    session_file.write_text(results_folder)
    log_dir = Path(results_folder)

    # Write participant_info.txt for any other scripts that need it
    info_file = Path(__file__).parent / "participant_info.txt"
    with open(info_file, "w") as f:
        f.write(f"{participant_name},{CONDITION},{attempts},{ID},{delaytime}")

    # Launch subscriber automatically
    subscriber_path = Path(__file__).parent / "qtm_zmq_subscriber.py"
    print(f"🚀 Launching subscriber: {subscriber_path.name}")
    subprocess.Popen([sys.executable, str(subscriber_path)], cwd=str(Path(__file__).parent))

    # Send config to subscriber via ZMQ config socket
    zmq_config_socket = zmq_context.socket(zmq.PUB)
    zmq_config_socket.bind(f"tcp://*:{ZMQ_CONFIG_PORT}")
    time.sleep(2)  # Wait for subscriber to connect

    config_data = {
        "participant_name": participant_name,
        "participant_folder": results_folder,
        "conditions": CONDITION,
        "attempts": attempts,
        "ID": ID,
        "delaytime": delaytime,
        "W_VALUES": W_VALUES,
        "D_VALUES": D_VALUES,
        "TOTAL_TRIALS": TOTAL_TRIALS,
    }
    # Send config multiple times to ensure delivery (PUB/SUB can miss first messages)
    for _ in range(5):
        zmq_config_socket.send_string(f"config {json.dumps(config_data)}")
        time.sleep(0.1)
    print(f"📤 Config sent to subscriber")

    output_file = log_dir / f"{participant_name}_{CONDITION}_ID{ID}_{attempts}_{delaytime}_touch_log.csv"
    clicked_file = log_dir / f"{participant_name}_{CONDITION}_ID{ID}_{attempts}_{delaytime}_clicked_log.csv"

    log_rows.clear()
    clicked_frames.clear()

    print(f"Vibration mode: {vibration_mode}")
    if vibration_mode == "motion-coupled":
        print(f"Output: {FREQUENCY}Hz, ±{AMPLITUDE}V sine wave (burst per bin change)")
    elif vibration_mode == "continuous":
        print(f"Output: {FREQUENCY}Hz, ±{AMPLITUDE}V sine wave (continuous while touching)")
    else:
        print(f"Output: NONE (no-vibration mode)")
    print(f"Stops after: {TOTAL_TRIALS} triggers")
    print(f"Bins: {NUM_BINS}, Range: {FSR_MIN}-{FSR_MAX}mm")
    print(f"Reference line: Left edge (bottom left to bottom right)")
    print("-" * 40)

    connection = await qtm_rt.connect(QTM_HOST)
    await connection.stream_frames(components=['3d', '6d'], on_packet=handle_qtm_data)

    print("🟢 Logging started...")

    # Wait for TOTAL_TRIALS triggers
    while trigger_count < TOTAL_TRIALS:
        await asyncio.sleep(0.05)
    
    streaming_enabled = False
    print(f"🛑 {TOTAL_TRIALS} triggers detected. Stopping...")

    print("Saving data...")

    # Save CSV
    headers = ['Frame', 'Pen X', 'Pen Y', 'Pen Z', 'Distance to Plane (mm)', 
               'Local X', 'Clicked']

    with open(output_file, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        for row in log_rows:
            frame = row[0]
            clicked = 1 if frame in clicked_frames else 0
            writer.writerow(row + [clicked])

    with open(clicked_file, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        for row in log_rows:
            frame = row[0]
            if frame in clicked_frames:
                writer.writerow(row + [1])

    print(f"✅ Data saved to: {output_file}")
    print(f"✅ Clicked data saved to: {clicked_file}")
    
    # Cleanup
    cleanup_daq()
    zmq_config_socket.close()
    zmq_socket.close()
    zmq_context.term()
    time.sleep(2)
    sys.exit(0)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nShutting down...")
        cleanup_daq()
    except Exception as e:
        print(f"Error: {e}")
        cleanup_daq()