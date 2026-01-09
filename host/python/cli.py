import argparse, sys, time
from datetime import datetime
from mss import mss
import numpy as np
import cv2
import serial
from serial.tools import list_ports

NUM_LEDS=60
TOP_LEDS=19
RIGHT_LEDS=11
BOTTOM_LEDS=19
LEFT_LEDS=11
RES=(128,128)

def find_port():
    ports=list_ports.comports()
    return ports[0].device if ports else None

def sample_perimeter(img):
    h,w,_=img.shape
    c=[]
    sw=w/TOP_LEDS
    for i in range(TOP_LEDS):
        x0=int(i*sw); x1=int((i+1)*sw); y1=max(1,int(h*0.12))
        b=img[0:y1,x0:x1]; a=b.reshape(-1,3).mean(0) if b.size else [0,0,0]; c.append(tuple(int(x) for x in a))
    sh=h/RIGHT_LEDS
    for i in range(RIGHT_LEDS):
        y0=int(i*sh); y1=int((i+1)*sh); x0=max(0,w-int(w*0.12))
        b=img[y0:y1,x0:w]; a=b.reshape(-1,3).mean(0) if b.size else [0,0,0]; c.append(tuple(int(x) for x in a))
    sw=w/BOTTOM_LEDS
    btm=[]
    for i in range(BOTTOM_LEDS):
        x0=int(i*sw); x1=int((i+1)*sw); y0=max(0,h-int(h*0.12))
        b=img[y0:h,x0:x1]; a=b.reshape(-1,3).mean(0) if b.size else [0,0,0]; btm.append(tuple(int(x) for x in a))
    btm.reverse(); c+=btm
    sh=h/LEFT_LEDS
    lft=[]
    for i in range(LEFT_LEDS):
        y0=int(i*sh); y1=int((i+1)*sh); x1=min(int(w*0.12),w)
        b=img[y0:y1,0:x1]; a=b.reshape(-1,3).mean(0) if b.size else [0,0,0]; lft.append(tuple(int(x) for x in a))
    lft.reverse(); c+=lft
    if len(c)<NUM_LEDS: c+=[(0,0,0)]*(NUM_LEDS-len(c))
    return c[:NUM_LEDS]

def formatted_now():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]

def main():
    p=argparse.ArgumentParser()
    p.add_argument('--port', '-p', default=None)
    p.add_argument('--baud', '-b', type=int, default=115200)
    p.add_argument('--fps', type=float, default=15.0)
    p.add_argument('--noblur', action='store_true')
    p.add_argument('--verbose', '-v', action='store_true')
    args=p.parse_args()
    port=args.port or find_port()
    if not port:
        print(f"{formatted_now()} No COM port found. Use --port to specify.")
        sys.exit(1)
    try:
        ser=serial.Serial(port, args.baud, timeout=1)
    except Exception as e:
        print(f"{formatted_now()} Failed to open serial: {e}")
        sys.exit(1)
    interval=1.0/args.fps
    sct=None
    try:
        sct=mss()
        monitor=sct.monitors[1]
        next_frame = time.perf_counter()
        while True:
            now = time.perf_counter()
            if now < next_frame:
                time.sleep(next_frame - now)
                now = time.perf_counter()
            t_frame_start = now
            s=sct.grab(monitor)
            img=np.array(s)[:,:,:3]
            img=cv2.cvtColor(img,cv2.COLOR_BGR2RGB)
            small=cv2.resize(img,RES,interpolation=cv2.INTER_AREA)
            if not args.noblur:
                small=cv2.GaussianBlur(small,(3,3),0)
            colors=sample_perimeter(small)
            data=bytearray()
            for r,g,b in colors:
                data.extend([r&0xFF,g&0xFF,b&0xFF])
            try:
                ser.write(data)
            except Exception as e:
                if args.verbose:
                    print(f"{formatted_now()} Serial write error: {e}")
            if args.verbose:
                elapsed_ms = (time.perf_counter() - t_frame_start) * 1000.0
                print(f"{formatted_now()} frame time {elapsed_ms:.1f} ms")
            next_frame += interval
            if next_frame < time.perf_counter():
                next_frame = time.perf_counter() + interval
    except KeyboardInterrupt:
        if args.verbose:
            print(f"{formatted_now()} Stopping on keyboard interrupt")
    except Exception as e:
        print(f"{formatted_now()} Error: {e}")
    finally:
        try:
            if sct: sct.close()
        except: pass
        try:
            ser.close()
        except: pass

if __name__=='__main__':
    main()
