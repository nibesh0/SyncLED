import serial
import time
import threading
import numpy as np
import cv2
from mss import mss
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import ttk
from serial.tools import list_ports
import colorsys
import psutil

try:
    import pynvml
    pynvml.nvmlInit()
    GPU_AVAILABLE = True
except Exception:
    GPU_AVAILABLE = False

try:
    import sounddevice as sd
    AUDIO_AVAILABLE = True
except Exception:
    AUDIO_AVAILABLE = False

NUM_LEDS = 96
TOP_LEDS = 31
RIGHT_LEDS = 17
BOTTOM_LEDS = 31
LEFT_LEDS = 17
RES = (128, 128)
FPS = 15
BYTE_TIMEOUT = 0.5
ACK_TIMEOUT = 0.25
MAX_RETRIES = 2

_prev_net = None
_net_lock = threading.Lock()

class Ambilight:
    def __init__(self, root):
        self.root = root
        self.running = False
        self.ser = None
        self.sct = None
        self.frame_id = 0
        self.photo = None
        self.led_rects = []
        self.audio_levels = np.zeros(NUM_LEDS, dtype=float)
        self.audio_lock = threading.Lock()
        self.audio_stream = None
        self.canvas_w = 640
        self.canvas_h = 460
        self.stats_height = 100  # Height reserved for stats at top
        self.last_stats_time = 0.0
        self.stats = {"time": "", "cpu": 0, "ram": 0, "gpu0": 0, "gpu1": 0, "dl": 0, "ul": 0, "cpu_ghz": 0.0}
        # History for graphs (last 60 samples)
        self.history_len = 60
        self.cpu_history = [0] * self.history_len
        self.ram_history = [0] * self.history_len
        self.gpu0_history = [0] * self.history_len
        self.gpu1_history = [0] * self.history_len
        self.setup_ui()
        if AUDIO_AVAILABLE:
            threading.Thread(target=self.start_audio_stream, daemon=True).start()
        self.root.after(1000, self.stats_tick)

    def list_ports(self):
        ports = list_ports.comports()
        return [p.device for p in ports]

    def setup_ui(self):
        self.canvas = tk.Canvas(self.root, width=self.canvas_w, height=self.canvas_h, bg="black")
        self.canvas.pack()
        f = tk.Frame(self.root)
        f.pack(fill='x')
        tk.Label(f, text="COM:").grid(row=0, column=0, padx=(4,0))
        self.port_var = tk.StringVar()
        ports = self.list_ports()
        self.port_combo = ttk.Combobox(f, textvariable=self.port_var, values=ports, width=12)
        if ports:
            self.port_combo.set(ports[0])
        else:
            self.port_combo.set("COM1")
        self.port_combo.grid(row=0, column=1)
        tk.Button(f, text="Refresh", command=self.refresh_ports).grid(row=0, column=2, padx=4)
        self.btn = tk.Button(f, text="Start", command=self.toggle, width=8)
        self.btn.grid(row=0, column=3, padx=8)
        self.status = tk.Label(f, text="Stopped", width=28, anchor='w')
        self.status.grid(row=0, column=4, padx=4)
        self.sens_var = tk.DoubleVar(value=1.0)
        tk.Label(f, text="Sens:").grid(row=0, column=5, padx=(10,0))
        tk.Scale(f, from_=0.0, to=5.0, resolution=0.1, orient='horizontal', variable=self.sens_var, length=120).grid(row=0, column=6, padx=4)
        self.create_led_rects()

    def refresh_ports(self):
        ports = self.list_ports()
        self.port_combo['values'] = ports
        if ports:
            self.port_combo.set(ports[0])

    def start_audio_stream(self):
        if not AUDIO_AVAILABLE:
            return
        if self.audio_stream:
            return
        try:
            FFT_SIZE = 2048
            def callback(indata, frames, time_info, status):
                data = indata.copy()
                if data.ndim > 1:
                    mono = data.mean(axis=1)
                else:
                    mono = data
                window = np.hanning(len(mono))
                spec = np.fft.rfft(mono * window, n=FFT_SIZE)
                mags = np.abs(spec)
                groups = np.array_split(mags, NUM_LEDS)
                energies = np.array([g.mean() if g.size else 0.0 for g in groups])
                maxv = energies.max() if energies.size else 1.0
                norm = energies / maxv if maxv >= 1e-9 else energies * 0.0
                with self.audio_lock:
                    self.audio_levels = norm
            self.audio_stream = sd.InputStream(callback=callback, channels=1, samplerate=44100, blocksize=1024)
            self.audio_stream.start()
        except Exception:
            self.audio_stream = None

    def stop_audio_stream(self):
        try:
            if self.audio_stream:
                self.audio_stream.stop()
                self.audio_stream.close()
        except Exception:
            pass
        self.audio_stream = None
        with self.audio_lock:
            self.audio_levels = np.zeros(NUM_LEDS, dtype=float)

    def toggle(self):
        if not self.running:
            port = self.port_var.get()
            try:
                self.ser = serial.Serial(port, 115200, timeout=BYTE_TIMEOUT)
                self.running = True
                self.btn.configure(text="Stop")
                self.status.configure(text=f"Running on {port}")
                threading.Thread(target=self.loop, daemon=True).start()
            except Exception as e:
                self.ser = None
                self.running = False
                self.btn.configure(text="Start")
                self.status.configure(text=f"Error opening port: {repr(e)}")
        else:
            self.running = False
            self.btn.configure(text="Start")
            self.status.configure(text="Stopped")
            if self.ser:
                try:
                    self.ser.close()
                except Exception:
                    pass
                self.ser = None

    def loop(self):
        dt = 1.0 / FPS
        try:
            self.sct = mss()
            monitor = self.sct.monitors[1]
            while self.running:
                t0 = time.time()
                img = self.capture_with_sct(self.sct, monitor)
                if img is not None:
                    colors = self.sample(img)
                    colors = self.apply_audio_to_colors(colors)
                    self.root.after(0, self.update_gui, img, colors)
                    self.send(colors)
                d = dt - (time.time() - t0)
                if d > 0:
                    time.sleep(d)
        except Exception:
            pass
        finally:
            if self.sct:
                try:
                    self.sct.close()
                except Exception:
                    pass
            self.running = False
            self.root.after(0, self.status.configure, {"text": "Stopped"})

    def capture_with_sct(self, sct, monitor):
        try:
            s = sct.grab(monitor)
            img = np.array(s)[:, :, :3]
            img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            small = cv2.resize(img, RES, interpolation=cv2.INTER_AREA)
            return small
        except Exception:
            return None

    def show(self, img):
        edge_thickness = 28
        left_edge = right_edge = top_edge = bottom_edge = edge_thickness
        # Offset for stats area at top
        stats_offset = self.stats_height
        cx0 = left_edge
        cy0 = stats_offset + top_edge
        cx1 = self.canvas_w - right_edge
        cy1 = self.canvas_h - bottom_edge
        cw = max(1, int(cx1 - cx0))
        ch = max(1, int(cy1 - cy0))
        d = cv2.resize(img, (cw, ch), interpolation=cv2.INTER_AREA)
        pil = Image.fromarray(d)
        self.photo = ImageTk.PhotoImage(image=pil)
        self.canvas.delete('bg')
        self.canvas.create_image(cx0, cy0, anchor='nw', image=self.photo, tag='bg')
        self.canvas.image = self.photo

    def create_led_rects(self):
        pad = 2
        edge_thickness = 28
        left_edge = right_edge = top_edge = bottom_edge = edge_thickness
        # Offset for stats area at top
        stats_offset = self.stats_height
        cx0 = left_edge
        cy0 = stats_offset + top_edge
        cx1 = self.canvas_w - right_edge
        cy1 = self.canvas_h - bottom_edge
        top_seg_w = (cx1 - cx0) / TOP_LEDS
        for i in range(TOP_LEDS):
            x0 = int(cx0 + i * top_seg_w) + pad
            x1 = int(cx0 + (i + 1) * top_seg_w) - pad
            y0 = stats_offset + pad
            y1 = stats_offset + top_edge - pad
            rid = self.canvas.create_rectangle(x0, y0, x1, y1, fill="#000000", outline="")
            self.led_rects.append(rid)
        right_seg_h = (cy1 - cy0) / RIGHT_LEDS
        for i in range(RIGHT_LEDS):
            y0 = int(cy0 + i * right_seg_h) + pad
            y1 = int(cy0 + (i + 1) * right_seg_h) - pad
            x0 = self.canvas_w - right_edge + pad
            x1 = self.canvas_w - pad
            rid = self.canvas.create_rectangle(x0, y0, x1, y1, fill="#000000", outline="")
            self.led_rects.append(rid)
        bottom_seg_w = (cx1 - cx0) / BOTTOM_LEDS
        for i in range(BOTTOM_LEDS):
            x0 = int(cx1 - (i + 1) * bottom_seg_w) + pad
            x1 = int(cx1 - i * bottom_seg_w) - pad
            y0 = self.canvas_h - bottom_edge + pad
            y1 = self.canvas_h - pad
            rid = self.canvas.create_rectangle(x0, y0, x1, y1, fill="#000000", outline="")
            self.led_rects.append(rid)
        left_seg_h = (cy1 - cy0) / LEFT_LEDS
        for i in range(LEFT_LEDS):
            y0 = int(cy1 - (i + 1) * left_seg_h) + pad
            y1 = int(cy1 - i * left_seg_h) - pad
            x0 = pad
            x1 = left_edge - pad
            rid = self.canvas.create_rectangle(x0, y0, x1, y1, fill="#000000", outline="")
            self.led_rects.append(rid)

    def update_led_rects(self, colors):
        for i, col in enumerate(colors[:NUM_LEDS]):
            r, g, b = col
            hexc = "#%02x%02x%02x" % (r, g, b)
            try:
                self.canvas.itemconfig(self.led_rects[i], fill=hexc)
            except Exception:
                pass

    def update_gui(self, img, colors):
        self.show(img)
        self.update_led_rects(colors)

    def sample(self, img):
        h, w, _ = img.shape
        c = []
        sw = w / TOP_LEDS
        for i in range(TOP_LEDS):
            x0 = int(i * sw)
            x1 = int((i + 1) * sw)
            y1 = max(1, int(h * 0.12))
            b = img[0:y1, x0:x1]
            a = b.reshape(-1, 3).mean(0) if b.size else [0, 0, 0]
            c.append(tuple(int(x) for x in a))
        sh = h / RIGHT_LEDS
        for i in range(RIGHT_LEDS):
            y0 = int(i * sh)
            y1 = int((i + 1) * sh)
            x0 = max(0, w - int(w * 0.12))
            b = img[y0:y1, x0:w]
            a = b.reshape(-1, 3).mean(0) if b.size else [0, 0, 0]
            c.append(tuple(int(x) for x in a))
        sw = w / BOTTOM_LEDS
        btm = []
        for i in range(BOTTOM_LEDS):
            x0 = int(i * sw)
            x1 = int((i + 1) * sw)
            y0 = max(0, h - int(h * 0.12))
            b = img[y0:h, x0:x1]
            a = b.reshape(-1, 3).mean(0) if b.size else [0, 0, 0]
            btm.append(tuple(int(x) for x in a))
        btm.reverse()
        c += btm
        sh = h / LEFT_LEDS
        lft = []
        for i in range(LEFT_LEDS):
            y0 = int(i * sh)
            y1 = int((i + 1) * sh)
            x1 = min(int(w * 0.12), w)
            b = img[y0:y1, 0:x1]
            a = b.reshape(-1, 3).mean(0) if b.size else [0, 0, 0]
            lft.append(tuple(int(x) for x in a))
        lft.reverse()
        c += lft
        if len(c) < NUM_LEDS:
            c += [(0, 0, 0)] * (NUM_LEDS - len(c))
        return c[:NUM_LEDS]

    def collect_stats(self):
        now = time.strftime("%Y-%m-%d %H:%M:%S")
        cpu = int(psutil.cpu_percent(interval=None))
        ram = int(psutil.virtual_memory().percent)
        gpu0 = gpu1 = 0
        if GPU_AVAILABLE:
            try:
                device_count = pynvml.nvmlDeviceGetCount()
                if device_count >= 1:
                    handle = pynvml.nvmlDeviceGetHandleByIndex(0)
                    util = pynvml.nvmlDeviceGetUtilizationRates(handle)
                    gpu0 = int(util.gpu)
                if device_count >= 2:
                    handle = pynvml.nvmlDeviceGetHandleByIndex(1)
                    util = pynvml.nvmlDeviceGetUtilizationRates(handle)
                    gpu1 = int(util.gpu)
            except Exception:
                gpu0 = gpu1 = 0
        net = psutil.net_io_counters()
        with _net_lock:
            global _prev_net
            if _prev_net is None:
                dl = ul = 0
                _prev_net = (net.bytes_recv, net.bytes_sent, time.time())
            else:
                prev_recv, prev_sent, prev_t = _prev_net
                now_t = time.time()
                dt = now_t - prev_t if now_t - prev_t > 0 else 1.0
                dl = int((net.bytes_recv - prev_recv) / 1024.0 / dt + 0.5)
                ul = int((net.bytes_sent - prev_sent) / 1024.0 / dt + 0.5)
                _prev_net = (net.bytes_recv, net.bytes_sent, now_t)
        # Update history
        self.cpu_history.pop(0)
        self.cpu_history.append(cpu)
        self.ram_history.pop(0)
        self.ram_history.append(ram)
        self.gpu0_history.pop(0)
        self.gpu0_history.append(gpu0)
        self.gpu1_history.pop(0)
        self.gpu1_history.append(gpu1)
        return {"time": now, "cpu": cpu, "ram": ram, "gpu0": gpu0, "gpu1": gpu1, "dl": dl/1024, "ul": ul/1024}

    def render_status_overlay(self, stats):
        x = 6
        y = 4
        lh = 18
        bar_w = 150
        bar_h = 12
        self.canvas.delete("status")
        self.canvas.delete("status_bg")
        # Stats background fills the stats_height area at top
        self.canvas.create_rectangle(0, 0, self.canvas_w, self.stats_height, fill="black", outline="", tag="status_bg")
        # CPU
        cpu_label = f"CPU {stats['cpu']}%"
        self.canvas.create_text(x, y, anchor="nw", fill="white", font=("Consolas", 9), text=cpu_label, tag="status")
        self.draw_bar(x + 80, y, bar_w, bar_h, stats['cpu'], "#00ff00")
        y += lh
        # RAM
        ram_label = f"RAM {stats['ram']}%"
        self.canvas.create_text(x, y, anchor="nw", fill="white", font=("Consolas", 9), text=ram_label, tag="status")
        self.draw_bar(x + 80, y, bar_w, bar_h, stats['ram'], "#00aaff")
        y += lh
        # GPU0 and GPU1
        gpu_label = f"G0 {stats['gpu0']}%"
        self.canvas.create_text(x, y, anchor="nw", fill="white", font=("Consolas", 9), text=gpu_label, tag="status")
        self.draw_bar(x + 70, y, 100, bar_h, stats['gpu0'], "#ff6600")
        gpu1_label = f"G1 {stats['gpu1']}%"
        self.canvas.create_text(x + 200, y, anchor="nw", fill="white", font=("Consolas", 9), text=gpu1_label, tag="status")
        self.draw_bar(x + 270, y, 100, bar_h, stats['gpu1'], "#ff00ff")
        y += lh
        # Network
        self.canvas.create_text(x, y, anchor="nw", fill="white", font=("Consolas", 10), text=f"D {stats['dl']}MB/s  U {stats['ul']}MB/s", tag="status")
        y += lh
        # Date/Time (last row)
        self.canvas.create_text(x, y, anchor="nw", fill="#888888", font=("Consolas", 9), text=stats["time"], tag="status")
        self.canvas.tag_raise("status")

    def draw_bar(self, x, y, w, h, pct, color):
        """Draw a filled bar showing percentage"""
        self.canvas.create_rectangle(x, y, x + w, y + h, outline="#444444", tag="status")
        fw = int((w - 2) * max(0, min(100, pct)) / 100)
        if fw > 0:
            self.canvas.create_rectangle(x + 1, y + 1, x + 1 + fw, y + h - 1, fill=color, outline="", tag="status")

    def stats_tick(self):
        self.stats = self.collect_stats()
        self.root.after(0, self.render_status_overlay, self.stats)
        # send status packet if serial open (rate: 1s)
        if self.ser and getattr(self.ser, "is_open", False):
            try:
                csv = f"{self.stats['time']},{self.stats['cpu']},{self.stats['ram']},{self.stats['gpu0']},{self.stats['gpu1']},{self.stats['dl']},{self.stats['ul']}"
                data = csv.encode('utf-8')[:240]
                L = len(data) & 0xFF
                s2 = 0x56
                for b in data:
                    s2 = (s2 + int(b)) & 0xFFFF
                chk2 = s2 & 0xFF
                pkt2 = bytearray([0xAA, 0x56, L]) + data + bytearray([chk2])
                self.ser.write(pkt2)
            except Exception:
                pass
        self.last_stats_time = time.time()
        self.root.after(1000, self.stats_tick)

    def enhance_colors(self, colors):
        out = []
        sat_boost = 1.35
        contrast = 1.12
        highlight_thresh = 200.0
        highlight_reduce_strength = 0.75
        gamma = 1.06
        for r, g, b in colors[:NUM_LEDS]:
            rf, gf, bf = r / 255.0, g / 255.0, b / 255.0
            lum = 0.2126 * rf + 0.7152 * gf + 0.0722 * bf
            if lum > (highlight_thresh / 255.0):
                excess = (lum - (highlight_thresh / 255.0)) / (1.0 - (highlight_thresh / 255.0))
                factor = 1.0 - excess * highlight_reduce_strength
                factor = max(0.25, factor)
                rf *= factor
                gf *= factor
                bf *= factor
            h, s, v = colorsys.rgb_to_hsv(rf, gf, bf)
            s = min(1.0, s * sat_boost)
            v = 0.5 + contrast * (v - 0.5)
            v = max(0.0, min(1.0, v * 0.98))
            rr, gg, bb = colorsys.hsv_to_rgb(h, s, v)
            rr = pow(max(0.0, min(1.0, rr)), gamma)
            gg = pow(max(0.0, min(1.0, gg)), gamma)
            bb = pow(max(0.0, min(1.0, bb)), gamma)
            out.append((int(rr * 255), int(gg * 255), int(bb * 255)))
        if len(out) < NUM_LEDS:
            out += [(0, 0, 0)] * (NUM_LEDS - len(out))
        return out[:NUM_LEDS]

    def apply_audio_to_colors(self, colors):
        colors2 = colors
        if AUDIO_AVAILABLE:
            sens = float(self.sens_var.get())
            with self.audio_lock:
                levels = self.audio_levels.copy() if hasattr(self, 'audio_levels') else np.zeros(NUM_LEDS, dtype=float)
            out = []
            for i, (r, g, b) in enumerate(colors[:NUM_LEDS]):
                lvl = float(levels[i]) if i < len(levels) else 0.0
                scale = 1.0 + sens * lvl
                rr = min(255, int(r * scale))
                gg = min(255, int(g * scale))
                bb = min(255, int(b * scale))
                out.append((rr, gg, bb))
            colors2 = out[:NUM_LEDS]
        colors2 = self.enhance_colors(colors2)
        return colors2[:NUM_LEDS]

    def send(self, c):
        if self.ser is None or not getattr(self.ser, "is_open", False):
            return
        payload = bytearray()
        for i in range(NUM_LEDS):
            r, g, b = c[i] if i < len(c) else (0, 0, 0)
            payload.extend([r & 0xFF, g & 0xFF, b & 0xFF])
        if len(payload) < NUM_LEDS * 3:
            payload.extend([0] * (NUM_LEDS * 3 - len(payload)))
        hdr = bytearray([0xAA, 0x55, self.frame_id & 0xFF])
        s = self.frame_id
        for x in payload:
            s = (s + int(x)) & 0xFFFF
        chk = s & 0xFF
        pkt = hdr + payload + bytearray([chk])
        try:
            self.ser.timeout = BYTE_TIMEOUT
            self.ser.write(pkt)
        except Exception:
            return
        got_ack = False
        for _ in range(MAX_RETRIES + 1):
            start = time.time()
            ack = b''
            while time.time() - start < ACK_TIMEOUT:
                try:
                    a = self.ser.read(1)
                    if a:
                        ack = a
                        break
                except Exception:
                    break
            if ack == b'A':
                got_ack = True
                break
            else:
                try:
                    self.ser.write(pkt)
                except Exception:
                    break
        if not got_ack:
            return
        self.frame_id = (self.frame_id + 1) & 0xFF

root = tk.Tk()
app = Ambilight(root)
root.mainloop()
