import serial, time, threading, numpy as np, cv2
from mss import mss
from PIL import Image, ImageTk
import tkinter as tk
from serial.tools import list_ports
import colorsys

try:
    import sounddevice as sd
    AUDIO_AVAILABLE = True
except Exception:
    AUDIO_AVAILABLE = False

NUM_LEDS=96
TOP_LEDS=31
RIGHT_LEDS=17
BOTTOM_LEDS=31
LEFT_LEDS=17
RES=(128,128)
FPS=30
APPLY_BLUR=True

class Ambilight:
    def __init__(self,root):
        self.root=root
        self.running=False
        self.ser=None
        self.sct=None
        self.monitor=None
        self.canvas_w=640
        self.canvas_h=360
        self.canvas=tk.Canvas(root,width=self.canvas_w,height=self.canvas_h,bg='black')
        self.canvas.pack()
        f=tk.Frame(root);f.pack(fill='x')
        tk.Label(f,text="COM:").grid(row=0,column=0)
        self.port_var=tk.StringVar(value=self.find_port())
        tk.Entry(f,textvariable=self.port_var,width=12).grid(row=0,column=1)
        self.btn=tk.Button(f,text="Start",command=self.toggle)
        self.btn.grid(row=0,column=2,padx=8)
        self.audio_var=tk.IntVar(value=1 if AUDIO_AVAILABLE else 0)
        self.audio_chk=tk.Checkbutton(f,text="Audio",variable=self.audio_var,command=self.on_audio_toggle)
        self.audio_chk.grid(row=0,column=3,padx=8)
        self.sens_var=tk.DoubleVar(value=1.0)
        tk.Label(f,text="Sens:").grid(row=0,column=4)
        tk.Scale(f,from_=0.0,to=5.0,resolution=0.1,orient='horizontal',variable=self.sens_var,length=120).grid(row=0,column=5)
        self.status=tk.Label(root,text="Stopped",anchor='w')
        self.status.pack(fill='x')
        self.photo=None
        self.led_rects=[]
        self.create_led_rects()
        self.audio_levels=np.zeros(NUM_LEDS,dtype=float)
        self.audio_lock=threading.Lock()
        self.audio_stream=None
        self.frame_id = 0
        if self.audio_var.get()==1 and AUDIO_AVAILABLE:
            threading.Thread(target=self.start_audio_stream,daemon=True).start()

    def find_port(self):
        ports=list_ports.comports()
        return ports[0].device if ports else "COM3"

    def create_led_rects(self):
        pad=2
        edge_thickness=28
        left_edge=right_edge=edge_thickness
        top_edge=bottom_edge=edge_thickness
        cx0=left_edge
        cy0=top_edge
        cx1=self.canvas_w-right_edge
        cy1=self.canvas_h-bottom_edge
        top_seg_w=(cx1-cx0)/TOP_LEDS
        for i in range(TOP_LEDS):
            x0=int(cx0 + i*top_seg_w)+pad
            x1=int(cx0 + (i+1)*top_seg_w)-pad
            y0=pad
            y1=top_edge-pad
            rid=self.canvas.create_rectangle(x0,y0,x1,y1,fill="#000000",outline="")
            self.led_rects.append(rid)
        right_seg_h=(cy1-cy0)/RIGHT_LEDS
        for i in range(RIGHT_LEDS):
            y0=int(cy0 + i*right_seg_h)+pad
            y1=int(cy0 + (i+1)*right_seg_h)-pad
            x0=self.canvas_w-right_edge+pad
            x1=self.canvas_w-pad
            rid=self.canvas.create_rectangle(x0,y0,x1,y1,fill="#000000",outline="")
            self.led_rects.append(rid)
        bottom_seg_w=(cx1-cx0)/BOTTOM_LEDS
        for i in range(BOTTOM_LEDS):
            x0=int(cx1 - (i+1)*bottom_seg_w)+pad
            x1=int(cx1 - i*bottom_seg_w)-pad
            y0=self.canvas_h-bottom_edge+pad
            y1=self.canvas_h-pad
            rid=self.canvas.create_rectangle(x0,y0,x1,y1,fill="#000000",outline="")
            self.led_rects.append(rid)
        left_seg_h=(cy1-cy0)/LEFT_LEDS
        for i in range(LEFT_LEDS):
            y0=int(cy1 - (i+1)*left_seg_h)+pad
            y1=int(cy1 - i*left_seg_h)-pad
            x0=pad
            x1=left_edge-pad
            rid=self.canvas.create_rectangle(x0,y0,x1,y1,fill="#000000",outline="")
            self.led_rects.append(rid)

    def on_audio_toggle(self):
        if self.audio_var.get()==1 and AUDIO_AVAILABLE:
            threading.Thread(target=self.start_audio_stream,daemon=True).start()
        else:
            self.stop_audio_stream()

    def start_audio_stream(self):
        if not AUDIO_AVAILABLE:
            self.root.after(0,self.status.configure,{"text":"Audio unavailable"})
            return
        if self.audio_stream:
            return
        try:
            FFT_SIZE=2048
            def callback(indata,frames,time_info,status):
                if status:
                    pass
                data = indata.copy()
                if data.ndim>1:
                    mono = data.mean(axis=1)
                else:
                    mono = data
                window = np.hanning(len(mono))
                spec = np.fft.rfft(mono*window,n=FFT_SIZE)
                mags = np.abs(spec)
                groups = np.array_split(mags, NUM_LEDS)
                energies = np.array([g.mean() if g.size else 0.0 for g in groups])
                maxv = energies.max() if energies.size else 1.0
                if maxv < 1e-9:
                    norm = energies*0.0
                else:
                    norm = energies / maxv
                with self.audio_lock:
                    self.audio_levels = norm
            self.audio_stream = sd.InputStream(callback=callback,channels=1,samplerate=44100,blocksize=1024)
            self.audio_stream.start()
            self.root.after(0,self.status.configure,{"text":"Audio running"})
        except Exception:
            self.audio_stream = None
            self.root.after(0,self.status.configure,{"text":"Audio error"})

    def stop_audio_stream(self):
        try:
            if self.audio_stream:
                self.audio_stream.stop()
                self.audio_stream.close()
        except Exception:
            pass
        self.audio_stream=None
        with self.audio_lock:
            self.audio_levels = np.zeros(NUM_LEDS,dtype=float)
        self.root.after(0,self.status.configure,{"text":"Audio stopped"})

    def toggle(self):
        if not self.running:
            try:
                self.ser=serial.Serial(self.port_var.get(),115200,timeout=0.05)
                self.running=True
                self.btn.configure(text="Stop")
                self.status.configure(text="Running")
                threading.Thread(target=self.loop,daemon=True).start()
            except Exception:
                self.status.configure(text="Error opening port")
        else:
            self.running=False
            self.btn.configure(text="Start")
            self.status.configure(text="Stopped")
            if self.ser:self.ser.close()

    def loop(self):
        dt=1.0/FPS
        sct=None
        try:
            sct=mss()
            monitor=sct.monitors[1]
            while self.running:
                t0=time.time()
                img=self.capture_with_sct(sct,monitor)
                if img is not None:
                    colors=self.sample(img)
                    colors=self.apply_audio_to_colors(colors)
                    self.root.after(0,self.update_gui,img,colors)
                    self.send(colors)
                d=dt-(time.time()-t0)
                if d>0:time.sleep(d)
        except Exception:
            pass
        finally:
            if sct:
                try:
                    sct.close()
                except Exception:
                    pass
            self.running=False
            self.root.after(0,self.status.configure,{"text":"Stopped"})

    def capture_with_sct(self,sct,monitor):
        try:
            s=sct.grab(monitor)
            img=np.array(s)[:,:,:3]
            img=cv2.cvtColor(img,cv2.COLOR_BGR2RGB)
            small=cv2.resize(img,RES,interpolation=cv2.INTER_AREA)
            if APPLY_BLUR:small=cv2.GaussianBlur(small,(3,3),0)
            return small
        except Exception:
            return None

    def update_gui(self,img,colors):
        self.show(img)
        self.update_led_rects(colors)

    def show(self,img):
        edge_thickness=28
        left_edge=right_edge=edge_thickness
        top_edge=bottom_edge=edge_thickness
        cx0=left_edge
        cy0=top_edge
        cx1=self.canvas_w-right_edge
        cy1=self.canvas_h-bottom_edge
        cw=max(1,int(cx1-cx0))
        ch=max(1,int(cy1-cy0))
        d=cv2.resize(img,(cw,ch),interpolation=cv2.INTER_AREA)
        pil=Image.fromarray(d)
        self.photo=ImageTk.PhotoImage(image=pil)
        self.canvas.create_image(cx0,cy0,anchor='nw',image=self.photo)

    def sample(self,img):
        h,w,_=img.shape
        c=[]
        sw=w/TOP_LEDS
        for i in range(TOP_LEDS):
            x0=int(i*sw);x1=int((i+1)*sw);y1=max(1,int(h*0.12))
            b=img[0:y1,x0:x1];a=b.reshape(-1,3).mean(0) if b.size else [0,0,0];c.append(tuple(int(x) for x in a))
        sh=h/RIGHT_LEDS
        for i in range(RIGHT_LEDS):
            y0=int(i*sh);y1=int((i+1)*sh);x0=max(0,w-int(w*0.12))
            b=img[y0:y1,x0:w];a=b.reshape(-1,3).mean(0) if b.size else [0,0,0];c.append(tuple(int(x) for x in a))
        sw=w/BOTTOM_LEDS
        btm=[]
        for i in range(BOTTOM_LEDS):
            x0=int(i*sw);x1=int((i+1)*sw);y0=max(0,h-int(h*0.12))
            b=img[y0:h,x0:x1];a=b.reshape(-1,3).mean(0) if b.size else [0,0,0];btm.append(tuple(int(x) for x in a))
        btm.reverse();c+=btm
        sh=h/LEFT_LEDS
        lft=[]
        for i in range(LEFT_LEDS):
            y0=int(i*sh);y1=int((i+1)*sh);x1=min(int(w*0.12),w)
            b=img[y0:y1,0:x1];a=b.reshape(-1,3).mean(0) if b.size else [0,0,0];lft.append(tuple(int(x) for x in a))
        lft.reverse();c+=lft
        if len(c)<NUM_LEDS:c+=[(0,0,0)]*(NUM_LEDS-len(c))
        return c[:NUM_LEDS]

    def enhance_colors(self,colors):
        out=[]
        sat_boost=1.35
        contrast=1.12
        highlight_thresh=200.0
        highlight_reduce_strength=0.75
        gamma=1.06
        for r,g,b in colors[:NUM_LEDS]:
            rf, gf, bf = r/255.0, g/255.0, b/255.0
            lum = 0.2126*rf + 0.7152*gf + 0.0722*bf
            if lum > (highlight_thresh/255.0):
                excess = (lum - (highlight_thresh/255.0)) / (1.0 - (highlight_thresh/255.0))
                factor = 1.0 - excess * highlight_reduce_strength
                factor = max(0.25, factor)
                rf *= factor; gf *= factor; bf *= factor
            h,s,v = colorsys.rgb_to_hsv(rf,gf,bf)
            s = min(1.0, s * sat_boost)
            v = 0.5 + contrast * (v - 0.5)
            v = max(0.0, min(1.0, v * 0.98))
            rr,gg,bb = colorsys.hsv_to_rgb(h,s,v)
            rr = pow(max(0.0,min(1.0,rr)), gamma)
            gg = pow(max(0.0,min(1.0,gg)), gamma)
            bb = pow(max(0.0,min(1.0,bb)), gamma)
            out.append((int(rr*255), int(gg*255), int(bb*255)))
        if len(out)<NUM_LEDS:
            out += [(0,0,0)]*(NUM_LEDS-len(out))
        return out[:NUM_LEDS]

    def apply_audio_to_colors(self,colors):
        colors2 = colors
        if AUDIO_AVAILABLE and self.audio_var.get()==1:
            sens = float(self.sens_var.get())
            with self.audio_lock:
                levels = self.audio_levels.copy() if hasattr(self,'audio_levels') else np.zeros(NUM_LEDS,dtype=float)
            out=[]
            for i,(r,g,b) in enumerate(colors[:NUM_LEDS]):
                lvl = float(levels[i]) if i < len(levels) else 0.0
                scale = 1.0 + sens * lvl
                rr = min(255,int(r*scale))
                gg = min(255,int(g*scale))
                bb = min(255,int(b*scale))
                out.append((rr,gg,bb))
            colors2 = out[:NUM_LEDS]
        colors2 = self.enhance_colors(colors2)
        return colors2[:NUM_LEDS]

    def send(self,c):
        if self.ser and self.ser.is_open:
            payload = bytearray()
            for r,g,b in c:
                payload.extend([r,g,b])
            hdr = bytearray([0xAA,0x55, self.frame_id & 0xFF])
            s = self.frame_id
            for x in payload: s += x
            chk = s & 0xFF
            packet = hdr + payload + bytearray([chk])
            retries = 1
            ack = False
            for _ in range(retries+1):
                try:
                    self.ser.write(packet)
                except Exception:
                    break
                try:
                    a = self.ser.read(1)
                    if a == b'A':
                        ack = True
                        break
                    else:
                        time.sleep(0.01)
                except Exception:
                    pass
            self.frame_id = (self.frame_id + 1) & 0xFF

    def update_led_rects(self,colors):
        for i,col in enumerate(colors[:NUM_LEDS]):
            r,g,b=col
            hexc="#%02x%02x%02x"%(r,g,b)
            try:
                self.canvas.itemconfig(self.led_rects[i],fill=hexc)
            except Exception:
                pass

root=tk.Tk()
app=Ambilight(root)
root.mainloop()
