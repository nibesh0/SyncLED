"""
Ambilight Color Wheel + Brightness + System Tray
- Click/drag wheel to choose color
- Brightness slider darkens/all LEDs and preview
- COM drop-down, Refresh, Start/Stop, Send to Serial (SIMPLE / FULL)
- Optional audio FFT boost (requires sounddevice)
- Tray integration using pystray: hide to tray, open, exit
"""

import threading, time, math, sys
import tkinter as tk
from tkinter import messagebox

# safe imports
try:
    from PIL import Image, ImageTk, ImageDraw
    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False

try:
    import numpy as np
except Exception:
    raise RuntimeError("numpy is required: pip install numpy")

try:
    import serial
    from serial.tools import list_ports
    PYSERIAL_AVAILABLE = True
except Exception:
    serial = None
    list_ports = None
    PYSERIAL_AVAILABLE = False

try:
    import sounddevice as sd
    AUDIO_AVAILABLE = True
except Exception:
    sd = None
    AUDIO_AVAILABLE = False

# tray
try:
    import pystray
    TRAY_AVAILABLE = True
except Exception:
    pystray = None
    TRAY_AVAILABLE = False

# configuration
NUM_LEDS = 96
TOP_LEDS = 31
RIGHT_LEDS = 17
BOTTOM_LEDS = 31
LEFT_LEDS = 17

class AmbiTrayApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Ambilight - Wheel + Brightness + Tray")
        self.base_color = (0,0,0)
        self.brightness = 100
        self.frame_id = 0
        self.ser = None
        self.running = False
        self.icon = None
        self.tray_thread = None
        self.tray_visible = False

        # top controls
        ctrl = tk.Frame(root)
        ctrl.pack(fill='x', padx=6, pady=6)

        tk.Label(ctrl, text="Port:").grid(row=0, column=0, sticky='w')
        self.port_var = tk.StringVar()
        self.port_menu = tk.OptionMenu(ctrl, self.port_var, "")
        self.port_menu.config(width=12)
        self.port_menu.grid(row=0, column=1, padx=(2,6))
        tk.Button(ctrl, text="Refresh", command=self.refresh_ports).grid(row=0, column=2, padx=(0,6))

        self.send_var = tk.IntVar(value=1 if PYSERIAL_AVAILABLE else 0)
        self.send_chk = tk.Checkbutton(ctrl, text="Send to Serial", variable=self.send_var)
        self.send_chk.grid(row=0, column=3, padx=(0,8))
        if not PYSERIAL_AVAILABLE:
            self.send_chk.config(state='disabled')

        tk.Label(ctrl, text="Packet:").grid(row=0, column=4, sticky='w')
        self.packet_var = tk.StringVar(value="FULL")
        tk.Radiobutton(ctrl, text="Full (AA 55 ...)", variable=self.packet_var, value="FULL").grid(row=0, column=5, padx=(0,6))
        tk.Radiobutton(ctrl, text="Simple (S R G B \\n)", variable=self.packet_var, value="SIMPLE").grid(row=0, column=6, padx=(0,6))

        self.start_btn = tk.Button(ctrl, text="Start", width=10, command=self.toggle_start)
        self.start_btn.grid(row=0, column=7, padx=(8,0))

        # audio controls
        self.audio_var = tk.IntVar(value=1 if AUDIO_AVAILABLE else 0)
        self.audio_chk = tk.Checkbutton(ctrl, text="Audio", variable=self.audio_var, command=self.on_audio_toggle)
        self.audio_chk.grid(row=0, column=8, padx=(10,0))
        if not AUDIO_AVAILABLE:
            self.audio_chk.config(state='disabled')
        tk.Label(ctrl, text="Sens:").grid(row=0, column=9, sticky='w', padx=(8,0))
        self.sens_var = tk.DoubleVar(value=1.0)
        tk.Scale(ctrl, from_=0.0, to=5.0, resolution=0.1, orient='horizontal', variable=self.sens_var, length=140).grid(row=0, column=10)

        # tray/hide button
        self.hide_btn = tk.Button(ctrl, text="Hide to Tray", width=12, command=self.hide_to_tray)
        self.hide_btn.grid(row=0, column=11, padx=(8,0))

        self.status = tk.Label(root, text="Stopped", anchor='w')
        self.status.pack(fill='x', padx=6)

        # canvas for leds
        self.canvas_w = 640
        self.canvas_h = 360
        self.canvas = tk.Canvas(root, width=self.canvas_w, height=self.canvas_h, bg='black')
        self.canvas.pack(side='left', padx=6, pady=6)

        # right side: wheel/preview/brightness
        right = tk.Frame(root)
        right.pack(side='left', fill='y', padx=6, pady=6)

        self.wheel_size = 300
        if PIL_AVAILABLE:
            self.wheel_img = self.make_color_wheel(self.wheel_size)
            self.wheel_photo = ImageTk.PhotoImage(self.wheel_img)
            self.wheel_canvas = tk.Canvas(right, width=self.wheel_size, height=self.wheel_size)
            self.wheel_canvas.pack(pady=(6,4))
            self.wheel_canvas.create_image(0,0,anchor='nw',image=self.wheel_photo)
            self.wheel_canvas.bind("<Button-1>", self.on_wheel)
            self.wheel_canvas.bind("<B1-Motion>", self.on_wheel)
        else:
            tk.Label(right, text="Pillow required for wheel").pack()

        self.preview = tk.Canvas(right, width=120, height=120, bg='#000000')
        self.preview.pack(pady=(8,4))
        self.hex_label = tk.Label(right, text="#000000", font=("Consolas", 12))
        self.hex_label.pack()

        # Brightness
        tk.Label(right, text="Brightness:").pack(pady=(10,0))
        self.brightness_var = tk.IntVar(value=100)
        self.brightness_slider = tk.Scale(right, from_=0, to=100, orient='horizontal',
                                          variable=self.brightness_var, length=220, command=self.on_brightness_change)
        self.brightness_slider.pack()

        # led rects
        self.led_rects = []
        self.create_led_rects()

        # audio internals
        self.audio_stream = None
        self.audio_levels = np.zeros(NUM_LEDS, dtype=float)
        self.audio_lock = threading.Lock()

        # setup tray icon if available
        if TRAY_AVAILABLE and PIL_AVAILABLE:
            self.create_tray_icon()
        else:
            if not TRAY_AVAILABLE:
                print("pystray not installed — tray functionality disabled.")
            if not PIL_AVAILABLE:
                print("Pillow not installed — tray icon disabled.")

        # intercept close to hide instead of exit
        self.root.protocol("WM_DELETE_WINDOW", self.hide_to_tray)

        # initial refresh ports
        self.refresh_ports()

    # ------------------- ports -------------------
    def refresh_ports(self):
        choices = []
        if PYSERIAL_AVAILABLE:
            ports = list_ports.comports()
            choices = [p.device for p in ports]
        if not choices:
            choices = ["COM3"]
        menu = self.port_menu["menu"]
        menu.delete(0, "end")
        for c in choices:
            menu.add_command(label=c, command=lambda v=c: self.port_var.set(v))
        if not self.port_var.get():
            self.port_var.set(choices[0])

    # ------------------- color wheel -------------------
    def make_color_wheel(self, size):
        img = Image.new('RGB', (size, size), (40,40,40))
        pix = img.load()
        rmax = size//2
        cx = cy = size//2
        for y in range(size):
            for x in range(size):
                dx = x-cx; dy = y-cy
                r = math.hypot(dx,dy)
                if r <= rmax:
                    angle = math.atan2(dy,dx)
                    h = (angle + math.pi)/(2*math.pi)
                    s = r / rmax
                    v = 1.0
                    rgb = self.hsv_to_rgb(h,s,v)
                    pix[x,y] = tuple(int(255*c) for c in rgb)
                else:
                    pix[x,y] = (40,40,40)
        return img

    def hsv_to_rgb(self,h,s,v):
        if s == 0.0: return (v,v,v)
        i = int(h*6.0)
        f = (h*6.0) - i
        p = v*(1.0 - s)
        q = v*(1.0 - s*f)
        t = v*(1.0 - s*(1.0 - f))
        i = i % 6
        if i==0: return (v,t,p)
        if i==1: return (q,v,p)
        if i==2: return (p,v,t)
        if i==3: return (p,q,v)
        if i==4: return (t,p,v)
        return (v,p,q)

    def on_wheel(self, event):
        x = event.x; y = event.y
        if 0 <= x < self.wheel_size and 0 <= y < self.wheel_size:
            rgb = self.wheel_img.getpixel((x,y))
            if rgb != (40,40,40):
                self.set_base_color(rgb)
                if self.running and self.send_var.get() and self.ser and getattr(self.ser,'is_open',False):
                    self.send_color_to_serial(self.get_scaled_color())

    def set_base_color(self, rgb):
        self.base_color = rgb
        scaled = self.get_scaled_color()
        hexc = '#%02x%02x%02x' % scaled
        self.preview.configure(bg=hexc)
        self.hex_label.configure(text=hexc)
        self.fill_leds(scaled)
        # optionally update tray icon small indicator (best-effort)
        if TRAY_AVAILABLE and self.icon:
            try:
                # create simple 32x32 icon with current color center
                img = Image.new('RGB', (32, 32), (30,30,30))
                d = ImageDraw.Draw(img)
                d.ellipse((4,4,28,28), fill=scaled)
                self.icon.icon = img
            except Exception:
                pass

    def get_scaled_color(self):
        b = self.brightness_var.get() / 100.0
        r = int(round(self.base_color[0] * b)) & 0xFF
        g = int(round(self.base_color[1] * b)) & 0xFF
        bl = int(round(self.base_color[2] * b)) & 0xFF
        return (r,g,bl)

    def on_brightness_change(self, _=None):
        scaled = self.get_scaled_color()
        hexc = '#%02x%02x%02x' % scaled
        self.preview.configure(bg=hexc)
        self.hex_label.configure(text=hexc)
        self.fill_leds(scaled)
        if self.running and self.send_var.get() and self.ser and getattr(self.ser,'is_open',False):
            self.send_color_to_serial(scaled)

    # ------------------- LEDs GUI -------------------
    def create_led_rects(self):
        pad=2
        edge_thickness=28
        left_edge=right_edge=edge_thickness
        top_edge=bottom_edge=edge_thickness
        cx0=left_edge; cy0=top_edge
        cx1=self.canvas_w-right_edge; cy1=self.canvas_h-bottom_edge
        # top
        top_seg_w=(cx1-cx0)/TOP_LEDS
        for i in range(TOP_LEDS):
            x0=int(cx0 + i*top_seg_w)+pad
            x1=int(cx0 + (i+1)*top_seg_w)-pad
            y0=pad; y1=top_edge-pad
            rid=self.canvas.create_rectangle(x0,y0,x1,y1,fill="#000000",outline="")
            self.led_rects.append(rid)
        # right
        right_seg_h=(cy1-cy0)/RIGHT_LEDS
        for i in range(RIGHT_LEDS):
            y0=int(cy0 + i*right_seg_h)+pad
            y1=int(cy0 + (i+1)*right_seg_h)-pad
            x0=self.canvas_w-right_edge+pad
            x1=self.canvas_w-pad
            rid=self.canvas.create_rectangle(x0,y0,x1,y1,fill="#000000",outline="")
            self.led_rects.append(rid)
        # bottom
        bottom_seg_w=(cx1-cx0)/BOTTOM_LEDS
        for i in range(BOTTOM_LEDS):
            x0=int(cx1 - (i+1)*bottom_seg_w)+pad
            x1=int(cx1 - i*bottom_seg_w)-pad
            y0=self.canvas_h-bottom_edge+pad
            y1=self.canvas_h-pad
            rid=self.canvas.create_rectangle(x0,y0,x1,y1,fill="#000000",outline="")
            self.led_rects.append(rid)
        # left
        left_seg_h=(cy1-cy0)/LEFT_LEDS
        for i in range(LEFT_LEDS):
            y0=int(cy1 - (i+1)*left_seg_h)+pad
            y1=int(cy1 - i*left_seg_h)-pad
            x0=pad; x1=left_edge-pad
            rid=self.canvas.create_rectangle(x0,y0,x1,y1,fill="#000000",outline="")
            self.led_rects.append(rid)

    def fill_leds(self, rgb):
        hexc = '#%02x%02x%02x' % rgb
        for rid in self.led_rects:
            try:
                self.canvas.itemconfig(rid, fill=hexc)
            except Exception:
                pass

    # ------------------- start/stop -------------------
    def toggle_start(self):
        if not self.running:
            if self.send_var.get() and not PYSERIAL_AVAILABLE:
                messagebox.showerror("pyserial missing", "pyserial not installed; can't send")
                return
            if self.send_var.get():
                port = self.port_var.get()
                try:
                    self.ser = serial.Serial(port, 115200, timeout=0.1)
                except Exception as e:
                    messagebox.showerror("Serial error", f"Cannot open {port}:\n{e}")
                    self.ser = None
                    return
            self.running = True
            self.start_btn.configure(text="Stop")
            self.status.configure(text="Running")
        else:
            self.running = False
            self.start_btn.configure(text="Start")
            self.status.configure(text="Stopped")
            if self.ser:
                try:
                    self.ser.close()
                except Exception:
                    pass
                self.ser = None

    # ------------------- serial sending -------------------
    def send_color_to_serial(self, rgb):
        if not (self.ser and getattr(self.ser,'is_open',False) and self.send_var.get()):
            return
        try:
            if self.packet_var.get() == "SIMPLE":
                pkt = b'S' + bytes([rgb[0]&0xFF, rgb[1]&0xFF, rgb[2]&0xFF]) + b'\n'
                self.ser.write(pkt)
            else:
                payload = bytearray()
                for _ in range(NUM_LEDS):
                    payload.extend([rgb[0]&0xFF, rgb[1]&0xFF, rgb[2]&0xFF])
                hdr = bytearray([0xAA, 0x55, self.frame_id & 0xFF])
                s = self.frame_id
                for x in payload:
                    s = (s + x) & 0xFFFFFFFF
                chk = s & 0xFF
                packet = hdr + payload + bytearray([chk])
                self.ser.write(packet)
                try:
                    _ = self.ser.read(1)
                except Exception:
                    pass
                self.frame_id = (self.frame_id + 1) & 0xFF
        except Exception as e:
            self.status.configure(text=f"Serial send error: {e}")
            try:
                self.ser.close()
            except Exception:
                pass
            self.ser = None
            self.send_var.set(0)

    # ------------------- audio (optional) -------------------
    def on_audio_toggle(self):
        if self.audio_var.get() and AUDIO_AVAILABLE:
            threading.Thread(target=self.start_audio_stream, daemon=True).start()
        else:
            self.stop_audio_stream()

    def start_audio_stream(self):
        if not AUDIO_AVAILABLE:
            self.status.configure(text="Audio unavailable")
            return
        if self.audio_stream: return
        try:
            FFT_SIZE = 2048
            def callback(indata, frames, time_info, status):
                if status: pass
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
                if maxv < 1e-9:
                    norm = energies * 0.0
                else:
                    norm = energies / maxv
                with self.audio_lock:
                    self.audio_levels = norm
            self.audio_stream = sd.InputStream(callback=callback, channels=1, samplerate=44100, blocksize=1024)
            self.audio_stream.start()
            self.status.configure(text="Audio running")
        except Exception as e:
            self.audio_stream = None
            self.status.configure(text=f"Audio error: {e}")

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
        self.status.configure(text="Audio stopped")

    # ------------------- tray integration -------------------
    def create_tray_icon(self):
        # create a simple default icon image (32x32)
        img = Image.new('RGB', (32, 32), (40,40,40))
        d = ImageDraw.Draw(img)
        d.ellipse((4,4,28,28), fill=(200,200,200))
        menu = pystray.Menu(
            pystray.MenuItem('Open Window', lambda _: self.show_window()),
            pystray.MenuItem('Hide to Tray', lambda _: self.hide_to_tray()),
            pystray.MenuItem('Exit', lambda _: self.exit_app())
        )
        self.icon = pystray.Icon("Ambilight", img, "Ambilight", menu)
        # run the tray icon in a background thread
        def run_icon():
            try:
                self.icon.run()
            except Exception:
                pass
        self.tray_thread = threading.Thread(target=run_icon, daemon=True)
        self.tray_thread.start()
        self.tray_visible = True

        # also set double-click to show window (some platforms)
        try:
            self.icon.visible = True
        except Exception:
            pass

    def show_window(self):
        # called from tray menu or icon double-click
        try:
            # Tk GUI ops must be called on main thread
            self.root.after(0, self._do_show)
        except Exception:
            self._do_show()

    def _do_show(self):
        self.root.deiconify()
        self.root.lift()
        try:
            self.root.focus_force()
        except Exception:
            pass

    def hide_to_tray(self):
        # hide GUI but keep app running; create tray icon if missing
        if TRAY_AVAILABLE and PIL_AVAILABLE and not self.icon:
            self.create_tray_icon()
        try:
            self.root.withdraw()
        except Exception:
            pass
        # status update
        self.status.configure(text="Hidden to tray (right-click icon to open/exit)")

    def exit_app(self):
        # stop tray icon and exit app
        try:
            if self.icon:
                try:
                    self.icon.stop()
                except Exception:
                    pass
            # close serial
            try:
                if self.ser:
                    self.ser.close()
            except Exception:
                pass
            # stop audio
            self.stop_audio_stream()
            # quit Tk
            self.root.after(0, self.root.quit)
        except Exception:
            try:
                self.root.quit()
            except Exception:
                pass

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("980x520")
    app = AmbiTrayApp(root)
    root.mainloop()
