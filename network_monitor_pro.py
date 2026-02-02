"""
========================================================
Network Monitor Pro (Windows)
========================================================
Author: You
Description:
Tkinter-based Windows network monitoring application
with ping tracking, live graph, alerts, tray icon,
airplane mode control, and logging.

Python Version: 3.10+
OS: Windows ONLY
========================================================
"""

# =========================
# Standard Library Imports
# =========================
import os
import sys
import csv
import time
import threading
import subprocess
import platform
from collections import deque
from datetime import datetime
import winsound

# =========================
# Third-Party Imports
# =========================
import tkinter as tk
from tkinter import ttk, messagebox

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

import pystray
from pystray import MenuItem
from PIL import Image, ImageDraw

# =========================
# Windows Check
# =========================
if platform.system() != "Windows":
    raise RuntimeError("This application runs on Windows only.")

# =========================
# Configuration
# =========================
PING_HOST = "8.8.8.8"
PING_INTERVAL = 1
ROLLING_BUFFER = 300        # 5 minutes
GRAPH_WINDOW = 60           # seconds
DISCONNECT_THRESHOLD = 4
LOG_FILE = "disconnect_log.csv"

# =========================
# Global State
# =========================
ping_history = deque(maxlen=ROLLING_BUFFER)
failure_streak = 0
total_pings = 0
successful_pings = 0
disconnect_counter = 0
airplane_mode = False
app_running = True
tray_icon = None

# =========================
# GUI Root (MUST BE FIRST)
# =========================
root = tk.Tk()
root.title("Network Monitor Pro")
root.geometry("900x600")

# =========================
# Tkinter Variables
# =========================
popup_enabled = tk.BooleanVar(root, value=True)
sound_enabled = tk.BooleanVar(root, value=True)
auto_airplane_enabled = tk.BooleanVar(root, value=True)

beep_frequency = tk.IntVar(root, value=800)

avg_ping_var = tk.StringVar(root, value="Avg Ping: -- ms")
fail_count_var = tk.StringVar(root, value="Failures (1m): 0")
uptime_var = tk.StringVar(root, value="Uptime: 0%")
startup_status_var = tk.StringVar(root, value="Startup: Disabled")

# =========================
# Utility Functions
# =========================
def run_powershell(cmd):
    subprocess.run(
        ["powershell", "-Command", cmd],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=subprocess.CREATE_NO_WINDOW
    )

# =========================
# Airplane Mode Control
# =========================
def set_airplane_mode(enable):
    global airplane_mode

    if enable:
        run_powershell(
            "Get-NetAdapter | Where-Object {$_.Status -eq 'Up'} | Disable-NetAdapter -Confirm:$false"
        )
        airplane_mode = True
    else:
        run_powershell(
            "Get-NetAdapter | Enable-NetAdapter -Confirm:$false"
        )
        airplane_mode = False

    update_airplane_label()

# =========================
# CSV Logging
# =========================
def log_disconnect():
    file_exists = os.path.exists(LOG_FILE)
    with open(LOG_FILE, "a", newline="") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["Timestamp"])
        writer.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S")])

# =========================
# Ping Thread
# =========================
def ping_loop():
    global failure_streak, total_pings, successful_pings, disconnect_counter

    while app_running:
        total_pings += 1
        latency = None

        try:
            output = subprocess.check_output(
                ["ping", "-n", "1", "-w", "1000", PING_HOST],
                universal_newlines=True,
                stderr=subprocess.DEVNULL
            )

            if "time=" in output:
                latency = int(output.split("time=")[1].split("ms")[0])
                successful_pings += 1
                failure_streak = 0
        except:
            failure_streak += 1

        ping_history.append(latency)

        # Disconnect detection
        if failure_streak >= DISCONNECT_THRESHOLD:
            disconnect_counter += 1

            if disconnect_counter % 4 == 0:
                log_disconnect()

            if popup_enabled.get():
                root.after(0, lambda: messagebox.showwarning(
                    "Network Disconnected",
                    "Internet connection lost!"
                ))

            if sound_enabled.get():
                winsound.Beep(beep_frequency.get(), 300)

            if auto_airplane_enabled.get():
                set_airplane_mode(True)
                time.sleep(3)
                set_airplane_mode(False)

        update_stats()
        time.sleep(PING_INTERVAL)

# =========================
# Stats Update
# =========================
def update_stats():
    last_minute = list(ping_history)[-60:]
    successes = [p for p in last_minute if p is not None]
    failures = last_minute.count(None)

    avg = int(sum(successes) / len(successes)) if successes else "--"
    avg_ping_var.set(f"Avg Ping: {avg} ms")
    fail_count_var.set(f"Failures (1m): {failures}")

    uptime = int((successful_pings / total_pings) * 100) if total_pings else 0
    uptime_var.set(f"Uptime: {uptime}%")

    update_tray_icon()

# =========================
# Status Detection
# =========================
def get_status():
    recent = list(ping_history)[-60:]
    if recent and all(p is None for p in recent):
        return "Disconnected"
    elif recent.count(None) > 4:
        return "Unstable"
    return "Online"

# =========================
# Tray Icon
# =========================
def make_icon(color):
    img = Image.new("RGB", (64, 64), color)
    d = ImageDraw.Draw(img)
    d.ellipse((8, 8, 56, 56), fill=color)
    return img

def update_tray_icon():
    if tray_icon:
        status = get_status()
        color = {"Online": "green", "Unstable": "yellow", "Disconnected": "red"}[status]
        tray_icon.icon = make_icon(color)

def tray_thread():
    global tray_icon

    def show():
        root.after(0, root.deiconify)

    def hide():
        root.after(0, root.withdraw)

    def quit_app():
        global app_running
        app_running = False
        tray_icon.stop()
        root.quit()

    tray_icon = pystray.Icon(
        "Network Monitor Pro",
        make_icon("green"),
        menu=pystray.Menu(
            MenuItem("Show", show),
            MenuItem("Hide", hide),
            MenuItem("Quit", quit_app)
        )
    )
    tray_icon.run()

# =========================
# GUI Layout
# =========================
top = ttk.Frame(root)
top.pack(fill="x", padx=10, pady=5)

ttk.Label(top, textvariable=avg_ping_var).pack(side="left", padx=10)
ttk.Label(top, textvariable=fail_count_var).pack(side="left", padx=10)
ttk.Label(top, textvariable=uptime_var).pack(side="left", padx=10)
ttk.Checkbutton(top, text="Popup Alerts", variable=popup_enabled).pack(side="right")

# Graph
fig = Figure(figsize=(8, 4))
ax = fig.add_subplot(111)
canvas = FigureCanvasTkAgg(fig, root)
canvas.get_tk_widget().pack(fill="both", expand=True)

def update_graph():
    ax.clear()
    data = list(ping_history)[-GRAPH_WINDOW:]

    for i, p in enumerate(data):
        if p is None:
            ax.axvline(i, color="red", linestyle="--", alpha=0.5)

    ax.plot([p if p is not None else 0 for p in data])
    ax.set_title("Live Ping (Last 60 Seconds)")
    canvas.draw()
    root.after(1000, update_graph)

# Bottom Controls
bottom = ttk.Frame(root)
bottom.pack(fill="x", padx=10, pady=5)

airplane_label = ttk.Label(bottom)
airplane_label.pack(side="left", padx=10)

def update_airplane_label():
    airplane_label.config(
        text=f"Airplane Mode: {'ON' if airplane_mode else 'OFF'}",
        foreground="red" if airplane_mode else "green"
    )

ttk.Button(bottom, text="Toggle Airplane",
           command=lambda: set_airplane_mode(not airplane_mode)).pack(side="left")

ttk.Checkbutton(bottom, text="Auto Airplane",
                variable=auto_airplane_enabled).pack(side="left", padx=10)

ttk.Checkbutton(bottom, text="Sound",
                variable=sound_enabled).pack(side="left")

ttk.Scale(bottom, from_=200, to=2000,
          variable=beep_frequency,
          orient="horizontal").pack(side="left", padx=10)

update_airplane_label()

# =========================
# Start Threads
# =========================
threading.Thread(target=ping_loop, daemon=True).start()
threading.Thread(target=tray_thread, daemon=True).start()

update_graph()

# =========================
# Start App
# =========================
root.mainloop()
