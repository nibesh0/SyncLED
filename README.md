# SyncLED

SyncLED is an ambient lighting system that synchronizes LED strips with your PC screen content. It supports both Python and C++ host applications and runs on an ESP32 microcontroller.

## Project Structure

```
/
├── firmware/
│   └── SyncLED/           # ESP32 Firmware (Arduino)
├── host/
│   ├── python/            # Python Host Applications
│   │   ├── gui.py         # GUI with Tray, Color Wheel & Settings
│   │   ├── cli.py         # Command Line Interface (lighter weight)
│   │   └── legacy/        # Legacy scripts
│   └── cpp/               # High-Performance C++ Host Applications
│       ├── console/       # Console-based C++ capture
│       └── tray/          # System Tray-based C++ capture
```

## Hardware Setup

1. **Components**:
   - ESP32 Development Board
   - WS2812B Addressable LED Strip
   - Power Supply (5V, Amperage depends on LED count)
   - Data connection (USB to ESP32)

2. **Wiring**:
   - Connect LED Strip Data Pin to **GPIO 5** on the ESP32 (Defined as `DATA_PIN` in `SyncLED.ino`).
   - Connect GND and 5V appropriately.

## Firmware Installation

1. Install the [Arduino IDE](https://www.arduino.cc/en/software).
2. Install the **FastLED** library via the Library Manager.
3. Open `firmware/SyncLED/SyncLED.ino`.
4. Configure the number of LEDs:
   ```cpp
   #define NUM_LEDS 96  // Update to your total LED count
   ```
5. Select your board (ESP32 Dev Module) and Port.
6. Upload the sketch.

## Host Software

### Python (Recommended for Ease of Use)

The Python implementation relies on `mss` for screen capture and `pyserial` for communication.

**Requirements:**
```bash
pip install mss pyserial pillow numpy sounddevice pystray
```
*(Note: `sounddevice` is optional for audio reactivity)*

**Running the GUI:**
The GUI version includes a color wheel, brightness control, and system tray integration.
```bash
python host/python/gui.py
```
- Select your COM port.
- Click "Start".
- Use the **Hide to Tray** button to keep it running in the background.

**Running the CLI:**
The CLI version is a lightweight console script.
```bash
python host/python/cli.py --port COM3 --baud 115200 --fps 30
```

### C++ (Recommended for Performance)

The C++ implementation uses DirectX 11 Desktop Duplication API for extremely low latency and low CPU usage.

**Compilation (Windows with Visual Studio):**
Open a "Developer Command Prompt for VS" and run:

*For Console Version:*
```cmd
cd host/cpp/console
cl /EHsc /O2 main.cpp /link d3d11.lib dxgi.lib
```

*For Tray Version:*
```cmd
cd host/cpp/tray
cl /EHsc /O2 main.cpp /link d3d11.lib dxgi.lib user32.lib shell32.lib comctl32.lib
```

**Usage:**
Run the compiled `.exe` files provided in the respective directories or move them to a convenient location.
