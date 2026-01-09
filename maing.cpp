// syncled_tray.cpp
// Build with Visual Studio:
// cl /EHsc /O2 syncled_tray.cpp /link d3d11.lib dxgi.lib user32.lib shell32.lib comctl32.lib

#define _CRT_SECURE_NO_WARNINGS
#include <windows.h>
#include <shellapi.h>
#include <d3d11.h>
#include <dxgi1_2.h>
#include <wrl/client.h>
#include <vector>
#include <array>
#include <chrono>
#include <thread>
#include <iostream>
#include <string>
#include <cstdint>
#include <algorithm>
#include <cmath>
#include <ctime>
#include <sstream>
#include <mutex>
#include <atomic>
#include <condition_variable>

using Microsoft::WRL::ComPtr;

// ------------------------------ your original constants ------------------------------
static const int DST_W = 128;
static const int DST_H = 128;
static const int NUM_LEDS = 96;
static const int TOP_LEDS = 31;
static const int RIGHT_LEDS = 17;
static const int BOTTOM_LEDS = 31;
static const int LEFT_LEDS = 17;

// ------------------------------ serial helper ------------------------------
HANDLE openSerial(const std::string& port, DWORD baud) {
    std::string full = "\\\\.\\" + port;
    HANDLE h = CreateFileA(full.c_str(), GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    if (h == INVALID_HANDLE_VALUE) return INVALID_HANDLE_VALUE;
    DCB dcb = {};
    dcb.DCBlength = sizeof(DCB);
    if (!GetCommState(h, &dcb)) { CloseHandle(h); return INVALID_HANDLE_VALUE; }
    dcb.BaudRate = baud;
    dcb.ByteSize = 8;
    dcb.Parity = NOPARITY;
    dcb.StopBits = ONESTOPBIT;
    if (!SetCommState(h, &dcb)) { CloseHandle(h); return INVALID_HANDLE_VALUE; }
    COMMTIMEOUTS to = {};
    to.ReadIntervalTimeout = 50;
    to.ReadTotalTimeoutConstant = 50;
    to.ReadTotalTimeoutMultiplier = 10;
    to.WriteTotalTimeoutConstant = 50;
    to.WriteTotalTimeoutMultiplier = 10;
    SetCommTimeouts(h, &to);
    return h;
}

// ------------------------------ color helpers ------------------------------
void rgb_to_hsv(float r, float g, float b, float &h, float &s, float &v) {
    float mx = std::max(r, std::max(g, b));
    float mn = std::min(r, std::min(g, b));
    v = mx;
    float d = mx - mn;
    s = (mx <= 0.0f) ? 0.0f : d / mx;
    if (d <= 1e-6f) { h = 0.0f; return; }
    if (mx == r) h = fmodf((g - b) / d, 6.0f);
    else if (mx == g) h = ((b - r) / d) + 2.0f;
    else h = ((r - g) / d) + 4.0f;
    h /= 6.0f;
    if (h < 0.0f) h += 1.0f;
}

void hsv_to_rgb(float h, float s, float v, float &r, float &g, float &b) {
    if (s <= 0.0f) { r = g = b = v; return; }
    h = h * 6.0f;
    int i = static_cast<int>(floor(h));
    float f = h - i;
    float p = v * (1.0f - s);
    float q = v * (1.0f - s * f);
    float t = v * (1.0f - s * (1.0f - f));
    switch (i % 6) {
    case 0: r = v; g = t; b = p; break;
    case 1: r = q; g = v; b = p; break;
    case 2: r = p; g = v; b = t; break;
    case 3: r = p; g = q; b = v; break;
    case 4: r = t; g = p; b = v; break;
    default: r = v; g = p; b = q; break;
    }
}

void enhance_packet(std::vector<uint8_t>& packet) {
    float sat_boost = 1.35f;
    float contrast = 1.12f;
    float highlight_thresh = 200.0f / 255.0f;
    float highlight_reduce_strength = 0.75f;
    float gamma = 1.06f;
    for (int i = 0; i < NUM_LEDS; ++i) {
        int idx = i * 3;
        float rf = packet[idx + 0] / 255.0f;
        float gf = packet[idx + 1] / 255.0f;
        float bf = packet[idx + 2] / 255.0f;
        float lum = 0.2126f * rf + 0.7152f * gf + 0.0722f * bf;
        if (lum > highlight_thresh) {
            float excess = (lum - highlight_thresh) / (1.0f - highlight_thresh);
            float factor = 1.0f - excess * highlight_reduce_strength;
            if (factor < 0.25f) factor = 0.25f;
            rf *= factor; gf *= factor; bf *= factor;
        }
        float h, s, v;
        rgb_to_hsv(rf, gf, bf, h, s, v);
        s = std::min(1.0f, s * sat_boost);
        v = 0.5f + contrast * (v - 0.5f);
        v = std::clamp(v * 0.98f, 0.0f, 1.0f);
        float rr, gg, bb;
        hsv_to_rgb(h, s, v, rr, gg, bb);
        rr = std::pow(std::clamp(rr, 0.0f, 1.0f), gamma);
        gg = std::pow(std::clamp(gg, 0.0f, 1.0f), gamma);
        bb = std::pow(std::clamp(bb, 0.0f, 1.0f), gamma);
        packet[idx + 0] = static_cast<uint8_t>(std::clamp(int(rr * 255.0f + 0.5f), 0, 255));
        packet[idx + 1] = static_cast<uint8_t>(std::clamp(int(gg * 255.0f + 0.5f), 0, 255));
        packet[idx + 2] = static_cast<uint8_t>(std::clamp(int(bb * 255.0f + 0.5f), 0, 255));
    }
}

std::vector<uint8_t> make_framed(const std::vector<uint8_t>& payload, uint8_t frame_id) {
    std::vector<uint8_t> out;
    out.reserve(3 + payload.size() + 1);
    out.push_back(0xAA);
    out.push_back(0x55);
    out.push_back(frame_id);
    uint32_t sum = frame_id;
    for (auto v : payload) { out.push_back(v); sum += v; }
    uint8_t chk = static_cast<uint8_t>(sum & 0xFF);
    out.push_back(chk);
    return out;
}

// ------------------------------ settings & synchronization ------------------------------
struct AppSettings {
    std::string port = "COM6";
    int baud = 115200;
    double fps = 30.0;
    bool use_header = true;
    bool expect_ack = true;
    bool once = true;
    bool test_glow_enable = false;
    int test_r = 128, test_g = 128, test_b = 128;
};

std::mutex settings_mtx;
AppSettings settings;

// thread control
std::atomic<bool> running(false);
std::atomic<bool> should_stop(false);
std::thread captureThread;
std::condition_variable_any capture_cv;

// forward declarations
void StartCaptureFromSettings();
void StopCaptureAndJoin();
void CaptureThreadFunc(AppSettings cfg);

// ------------------------------ tray UI / Win32 basics ------------------------------
const UINT WM_TRAYICON = WM_APP + 100;
const int ID_TRAY_APP_ICON = 5000;
const int ID_TRAY_STARTSTOP = 4001;
const int ID_TRAY_SETTINGS = 4002;
const int ID_TRAY_EXIT = 4003;
const int ID_TRAY_TOGGLE_HEADER = 4004;
const int ID_TRAY_TOGGLE_ACK = 4005;

HWND g_hwnd = nullptr;

void ShowTrayMenu(HWND hwnd) {
    POINT pt;
    GetCursorPos(&pt);
    HMENU hMenu = CreatePopupMenu();
    std::scoped_lock lock(settings_mtx);
    AppendMenu(hMenu, MF_STRING, ID_TRAY_STARTSTOP, running ? L"Stop Capture" : L"Start Capture");
    AppendMenu(hMenu, MF_STRING, ID_TRAY_SETTINGS, L"Settings...");
    AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenu(hMenu, MF_STRING | (settings.use_header ? MF_CHECKED : 0), ID_TRAY_TOGGLE_HEADER, L"Use Header");
    AppendMenu(hMenu, MF_STRING | (settings.expect_ack ? MF_CHECKED : 0), ID_TRAY_TOGGLE_ACK, L"Expect ACK");
    AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenu(hMenu, MF_STRING, ID_TRAY_EXIT, L"Exit");
    // ensure the window is foreground so menu closes correctly
    SetForegroundWindow(hwnd);
    TrackPopupMenu(hMenu, TPM_BOTTOMALIGN | TPM_LEFTALIGN, pt.x, pt.y, 0, hwnd, NULL);
    DestroyMenu(hMenu);
}

// --------- simple settings window (programmatic controls) ----------
INT_PTR CALLBACK SettingsProc(HWND hwndDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    static HWND hePort = nullptr, heBaud = nullptr, heFPS = nullptr, hbHeader = nullptr, hbAck = nullptr;
    static HWND heR = nullptr, heG = nullptr, heB = nullptr, hbTestEnable = nullptr, hbApply = nullptr, hbClose = nullptr;

    switch (message) {
    case WM_CREATE:
        return TRUE;
    case WM_INITDIALOG:
    {
        // create controls
        hePort = CreateWindowEx(0, L"EDIT", NULL, WS_CHILD | WS_VISIBLE | WS_BORDER | ES_LEFT,
            100, 20, 160, 20, hwndDlg, NULL, NULL, NULL);
        CreateWindowEx(0, L"STATIC", L"Port (eg. COM3):", WS_CHILD | WS_VISIBLE, 10, 20, 120, 20, hwndDlg, NULL, NULL, NULL);

        heBaud = CreateWindowEx(0, L"EDIT", NULL, WS_CHILD | WS_VISIBLE | WS_BORDER | ES_LEFT,
            100, 50, 160, 20, hwndDlg, NULL, NULL, NULL);
        CreateWindowEx(0, L"STATIC", L"Baud:", WS_CHILD | WS_VISIBLE, 10, 50, 120, 20, hwndDlg, NULL, NULL, NULL);

        heFPS = CreateWindowEx(0, L"EDIT", NULL, WS_CHILD | WS_VISIBLE | WS_BORDER | ES_LEFT,
            100, 80, 160, 20, hwndDlg, NULL, NULL, NULL);
        CreateWindowEx(0, L"STATIC", L"FPS:", WS_CHILD | WS_VISIBLE, 10, 80, 120, 20, hwndDlg, NULL, NULL, NULL);

        hbHeader = CreateWindowEx(0, L"BUTTON", L"Use Header", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
            10, 110, 120, 20, hwndDlg, NULL, NULL, NULL);

        hbAck = CreateWindowEx(0, L"BUTTON", L"Expect ACK", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
            140, 110, 120, 20, hwndDlg, NULL, NULL, NULL);

        CreateWindowEx(0, L"STATIC", L"Test Glow (R G B):", WS_CHILD | WS_VISIBLE, 10, 140, 140, 20, hwndDlg, NULL, NULL, NULL);
        heR = CreateWindowEx(0, L"EDIT", NULL, WS_CHILD | WS_VISIBLE | WS_BORDER | ES_LEFT, 150, 140, 40, 20, hwndDlg, NULL, NULL, NULL);
        heG = CreateWindowEx(0, L"EDIT", NULL, WS_CHILD | WS_VISIBLE | WS_BORDER | ES_LEFT, 200, 140, 40, 20, hwndDlg, NULL, NULL, NULL);
        heB = CreateWindowEx(0, L"EDIT", NULL, WS_CHILD | WS_VISIBLE | WS_BORDER | ES_LEFT, 250, 140, 40, 20, hwndDlg, NULL, NULL, NULL);
        hbTestEnable = CreateWindowEx(0, L"BUTTON", L"Enable Test Glow", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
            10, 170, 160, 20, hwndDlg, NULL, NULL, NULL);

        hbApply = CreateWindowEx(0, L"BUTTON", L"Apply", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
            60, 210, 80, 24, hwndDlg, (HMENU)1001, NULL, NULL);
        hbClose = CreateWindowEx(0, L"BUTTON", L"Close", WS_CHILD | WS_VISIBLE,
            180, 210, 80, 24, hwndDlg, (HMENU)1002, NULL, NULL);

        // populate from settings
        {
            std::scoped_lock lock(settings_mtx);
            SetWindowTextA(hePort, settings.port.c_str());
            {
                char buf[64];
                sprintf_s(buf, "%d", settings.baud);
                SetWindowTextA(heBaud, buf);
            }
            {
                char buf[64];
                sprintf_s(buf, "%.2f", settings.fps);
                SetWindowTextA(heFPS, buf);
            }
            SendMessage(hbHeader, BM_SETCHECK, settings.use_header ? BST_CHECKED : BST_UNCHECKED, 0);
            SendMessage(hbAck, BM_SETCHECK, settings.expect_ack ? BST_CHECKED : BST_UNCHECKED, 0);
            SendMessage(hbTestEnable, BM_SETCHECK, settings.test_glow_enable ? BST_CHECKED : BST_UNCHECKED, 0);
            {
                char buf[32];
                sprintf_s(buf, "%d", settings.test_r);
                SetWindowTextA(heR, buf);
                sprintf_s(buf, "%d", settings.test_g);
                SetWindowTextA(heG, buf);
                sprintf_s(buf, "%d", settings.test_b);
                SetWindowTextA(heB, buf);
            }
        }
    }
        return TRUE;
    case WM_COMMAND:
    {
        int id = LOWORD(wParam);
        if (id == 1001) { // Apply
            // read fields and update settings
            char buf[256];
            GetWindowTextA(GetDlgItem(hwndDlg, 0), buf, sizeof(buf)); // noop, but avoid warnings
            std::string newPort;
            char portbuf[128]; GetWindowTextA(GetDlgItem(hwndDlg, 0), portbuf, 128);
            // Actually we stored handles in static variables
            // easier: fetch edits we created via FindWindowEx by position (we preserved handles in static variables closures above)
            // But inside dialog proc we have them as static variables; use them:
            extern HWND hePort; // forward declare not ideal but we'll instead just capture via GetDlgItem approach is messy.
            // Simpler: we will fetch by enumerating child windows by type/position? To avoid complexity, store handles in window properties.
        }
        break;
    }
    case WM_CLOSE:
        DestroyWindow(hwndDlg);
        return TRUE;
    case WM_DESTROY:
        return TRUE;
    }
    return FALSE;
}

// Because making a fully-featured dialog with CreateDialogParam and resources in a single-file is verbose,
// we'll implement a very small custom settings window using CreateWindow and controls, not a dialogproc.
// To keep the code moderate in length, we'll create a function that creates a modeless settings window
// and processes its messages in the main message loop. This avoids dialog resources.

HWND settingsWnd = nullptr;
HWND editPort = nullptr, editBaud = nullptr, editFPS = nullptr, chkHeader = nullptr, chkAck = nullptr;
HWND chkTestEnable = nullptr, editR = nullptr, editG = nullptr, editB = nullptr;
HWND btnApply = nullptr, btnClose = nullptr;

LRESULT CALLBACK MainWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);

void CreateSettingsWindow(HWND parent) {
    if (settingsWnd && IsWindow(settingsWnd)) {
        SetForegroundWindow(settingsWnd);
        return;
    }
    const int W = 340, H = 270;
    settingsWnd = CreateWindowEx(WS_EX_WINDOWEDGE, L"STATIC", L"SyncLED Settings",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
        CW_USEDEFAULT, CW_USEDEFAULT, W, H, parent, NULL, GetModuleHandle(NULL), NULL);
    if (!settingsWnd) return;

    CreateWindowEx(0, L"STATIC", L"Port (e.g. COM3):", WS_CHILD | WS_VISIBLE, 12, 12, 120, 20, settingsWnd, NULL, NULL, NULL);
    editPort = CreateWindowEx(WS_EX_CLIENTEDGE, L"EDIT", NULL, WS_CHILD | WS_VISIBLE | ES_LEFT, 140, 12, 180, 20, settingsWnd, NULL, NULL, NULL);

    CreateWindowEx(0, L"STATIC", L"Baud:", WS_CHILD | WS_VISIBLE, 12, 42, 120, 20, settingsWnd, NULL, NULL, NULL);
    editBaud = CreateWindowEx(WS_EX_CLIENTEDGE, L"EDIT", NULL, WS_CHILD | WS_VISIBLE | ES_LEFT, 140, 42, 180, 20, settingsWnd, NULL, NULL, NULL);

    CreateWindowEx(0, L"STATIC", L"FPS:", WS_CHILD | WS_VISIBLE, 12, 72, 120, 20, settingsWnd, NULL, NULL, NULL);
    editFPS = CreateWindowEx(WS_EX_CLIENTEDGE, L"EDIT", NULL, WS_CHILD | WS_VISIBLE | ES_LEFT, 140, 72, 180, 20, settingsWnd, NULL, NULL, NULL);

    chkHeader = CreateWindowEx(0, L"BUTTON", L"Use Header", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 12, 102, 120, 22, settingsWnd, NULL, NULL, NULL);
    chkAck = CreateWindowEx(0, L"BUTTON", L"Expect ACK", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 140, 102, 120, 22, settingsWnd, NULL, NULL, NULL);

    CreateWindowEx(0, L"STATIC", L"Test Glow (R G B):", WS_CHILD | WS_VISIBLE, 12, 132, 130, 20, settingsWnd, NULL, NULL, NULL);
    editR = CreateWindowEx(WS_EX_CLIENTEDGE, L"EDIT", NULL, WS_CHILD | WS_VISIBLE | ES_LEFT, 140, 132, 40, 20, settingsWnd, NULL, NULL, NULL);
    editG = CreateWindowEx(WS_EX_CLIENTEDGE, L"EDIT", NULL, WS_CHILD | WS_VISIBLE | ES_LEFT, 190, 132, 40, 20, settingsWnd, NULL, NULL, NULL);
    editB = CreateWindowEx(WS_EX_CLIENTEDGE, L"EDIT", NULL, WS_CHILD | WS_VISIBLE | ES_LEFT, 240, 132, 40, 20, settingsWnd, NULL, NULL, NULL);
    chkTestEnable = CreateWindowEx(0, L"BUTTON", L"Enable Test Glow", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 12, 160, 160, 22, settingsWnd, NULL, NULL, NULL);

    btnApply = CreateWindowEx(0, L"BUTTON", L"Apply", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 60, 200, 80, 26, settingsWnd, (HMENU)3001, NULL, NULL);
    btnClose = CreateWindowEx(0, L"BUTTON", L"Close", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 180, 200, 80, 26, settingsWnd, (HMENU)3002, NULL, NULL);

    // populate values
    {
        std::scoped_lock lock(settings_mtx);
        SetWindowTextA(editPort, settings.port.c_str());
        char buf[64];
        sprintf_s(buf, "%d", settings.baud); SetWindowTextA(editBaud, buf);
        sprintf_s(buf, "%.2f", settings.fps); SetWindowTextA(editFPS, buf);
        SendMessage(chkHeader, BM_SETCHECK, settings.use_header ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessage(chkAck, BM_SETCHECK, settings.expect_ack ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessage(chkTestEnable, BM_SETCHECK, settings.test_glow_enable ? BST_CHECKED : BST_UNCHECKED, 0);
        sprintf_s(buf, "%d", settings.test_r); SetWindowTextA(editR, buf);
        sprintf_s(buf, "%d", settings.test_g); SetWindowTextA(editG, buf);
        sprintf_s(buf, "%d", settings.test_b); SetWindowTextA(editB, buf);
    }

    ShowWindow(settingsWnd, SW_SHOW);
}

void DestroySettingsWindow() {
    if (settingsWnd && IsWindow(settingsWnd)) {
        DestroyWindow(settingsWnd);
        settingsWnd = nullptr;
    }
}

// Utility to read integers from edit control safely
int ReadIntFromEdit(HWND h) {
    char buf[64] = {};
    if (!h) return 0;
    GetWindowTextA(h, buf, sizeof(buf));
    return atoi(buf);
}
double ReadDoubleFromEdit(HWND h) {
    char buf[64] = {};
    if (!h) return 0.0;
    GetWindowTextA(h, buf, sizeof(buf));
    return atof(buf);
}
std::string ReadStringFromEdit(HWND h) {
    char buf[256] = {};
    if (!h) return std::string();
    GetWindowTextA(h, buf, sizeof(buf));
    return std::string(buf);
}

// ------------------------------ capture thread implementation ------------------------------
void CaptureThreadFunc(AppSettings cfg) {
    // Note: This is adapted from your main loop. It creates its own D3D device & duplication.
    // Open serial
    if (cfg.test_glow_enable) {
        // quickly write the test glow then exit if once or just write and continue depending on cfg.once
        HANDLE hSerial = openSerial(cfg.port, (DWORD)cfg.baud);
        if (hSerial != INVALID_HANDLE_VALUE) {
            std::vector<uint8_t> payload(NUM_LEDS * 3);
            int rr = std::clamp(cfg.test_r, 0, 255);
            int gg = std::clamp(cfg.test_g, 0, 255);
            int bb = std::clamp(cfg.test_b, 0, 255);
            for (int i = 0; i < NUM_LEDS; ++i) {
                payload[i*3 + 0] = (uint8_t)rr;
                payload[i*3 + 1] = (uint8_t)gg;
                payload[i*3 + 2] = (uint8_t)bb;
            }
            std::vector<uint8_t> outbuf = cfg.use_header ? make_framed(payload, 0) : payload;
            DWORD written=0;
            WriteFile(hSerial, outbuf.data(), (DWORD)outbuf.size(), &written, NULL);
            CloseHandle(hSerial);
            if (cfg.once) {
                running = false;
                return;
            }
        } else {
            // can't open serial but continue to try capture
        }
    }

    UINT createFlags = D3D11_CREATE_DEVICE_BGRA_SUPPORT;
    ComPtr<ID3D11Device> device;
    ComPtr<ID3D11DeviceContext> context;
    D3D_FEATURE_LEVEL fl;
    if (FAILED(D3D11CreateDevice(nullptr, D3D_DRIVER_TYPE_HARDWARE, nullptr, createFlags, nullptr, 0, D3D11_SDK_VERSION, &device, &fl, &context))) {
        running = false;
        return;
    }
    ComPtr<IDXGIDevice> dxgiDevice;
    if (FAILED(device.As(&dxgiDevice))) { running = false; return; }
    ComPtr<IDXGIAdapter> adapter;
    if (FAILED(dxgiDevice->GetAdapter(&adapter))) { running = false; return; }
    ComPtr<IDXGIOutput> output;
    if (FAILED(adapter->EnumOutputs(0, &output))) { running = false; return; }
    ComPtr<IDXGIOutput1> output1;
    if (FAILED(output.As(&output1))) { running = false; return; }
    ComPtr<IDXGIOutputDuplication> deskDupl;
    if (FAILED(output1->DuplicateOutput(device.Get(), &deskDupl))) {
        running = false;
        return;
    }

    ComPtr<ID3D11Texture2D> staging;
    int src_w = 0, src_h = 0;
    std::vector<int> xs(DST_W + 1), ys(DST_H + 1);
    std::vector<uint8_t> dst(DST_W * DST_H * 3);
    std::vector<uint8_t> packet;
    packet.resize(NUM_LEDS * 3);

    std::chrono::duration<double> frameIntervalD(1.0 / cfg.fps);
    auto frameInterval = std::chrono::duration_cast<std::chrono::steady_clock::duration>(frameIntervalD);
    auto nextFrame = std::chrono::steady_clock::now();

    uint8_t frame_id = 0;

    // open serial (we open here and keep it for the loop)
    HANDLE hSerial = openSerial(cfg.port, (DWORD)cfg.baud);
    if (hSerial == INVALID_HANDLE_VALUE) {
        // we'll still attempt to capture frames but can't send them until serial available
    }

    while (running && !should_stop) {
        // check for settings changes by reading the current settings object (we capture a snapshot only if needed)
        // For simplicity this thread honors the initial cfg; if you want it to pick up live changes repeatedly,
        // you'd check a shared settings object periodically. Current design: stop/restart thread when settings change.

        nextFrame += frameInterval;

        ComPtr<IDXGIResource> desktopResource;
        DXGI_OUTDUPL_FRAME_INFO frameInfo;
        HRESULT hr = deskDupl->AcquireNextFrame(500, &frameInfo, &desktopResource);
        if (hr == DXGI_ERROR_WAIT_TIMEOUT) {
            std::this_thread::sleep_until(nextFrame);
            continue;
        }
        if (FAILED(hr)) break;

        ComPtr<ID3D11Texture2D> acquired;
        if (FAILED(desktopResource.As(&acquired))) {
            deskDupl->ReleaseFrame();
            std::this_thread::sleep_until(nextFrame);
            continue;
        }

        D3D11_TEXTURE2D_DESC srcDesc;
        acquired->GetDesc(&srcDesc);
        if (!staging || srcDesc.Width != (UINT)src_w || srcDesc.Height != (UINT)src_h) {
            src_w = (int)srcDesc.Width;
            src_h = (int)srcDesc.Height;
            D3D11_TEXTURE2D_DESC desc = {};
            desc.Width = srcDesc.Width;
            desc.Height = srcDesc.Height;
            desc.MipLevels = 1;
            desc.ArraySize = 1;
            desc.Format = srcDesc.Format;
            desc.SampleDesc.Count = 1;
            desc.Usage = D3D11_USAGE_STAGING;
            desc.BindFlags = 0;
            desc.CPUAccessFlags = D3D11_CPU_ACCESS_READ;
            desc.MiscFlags = 0;
            device->CreateTexture2D(&desc, nullptr, &staging);
            for (int i = 0; i <= DST_W; ++i) xs[i] = (src_w * i) / DST_W;
            for (int i = 0; i <= DST_H; ++i) ys[i] = (src_h * i) / DST_H;
        }

        context->CopyResource(staging.Get(), acquired.Get());
        D3D11_MAPPED_SUBRESOURCE mapped;
        if (FAILED(context->Map(staging.Get(), 0, D3D11_MAP_READ, 0, &mapped))) {
            deskDupl->ReleaseFrame();
            std::this_thread::sleep_until(nextFrame);
            continue;
        }

        uint8_t* srcPtr = reinterpret_cast<uint8_t*>(mapped.pData);
        int rowPitch = mapped.RowPitch;

        for (int dy = 0; dy < DST_H; ++dy) {
            int y0 = ys[dy];
            int y1 = ys[dy + 1];
            if (y1 <= y0) y1 = y0 + 1;
            for (int dx = 0; dx < DST_W; ++dx) {
                int x0 = xs[dx];
                int x1 = xs[dx + 1];
                if (x1 <= x0) x1 = x0 + 1;
                uint64_t rsum = 0, gsum = 0, bsum = 0;
                for (int sy = y0; sy < y1; ++sy) {
                    uint8_t* row = srcPtr + sy * rowPitch;
                    for (int sx = x0; sx < x1; ++sx) {
                        uint8_t b = row[sx * 4 + 0];
                        uint8_t g = row[sx * 4 + 1];
                        uint8_t r = row[sx * 4 + 2];
                        bsum += b; gsum += g; rsum += r;
                    }
                }
                int area = (x1 - x0) * (y1 - y0);
                int idx = (dy * DST_W + dx) * 3;
                dst[idx + 0] = (uint8_t)(rsum / area);
                dst[idx + 1] = (uint8_t)(gsum / area);
                dst[idx + 2] = (uint8_t)(bsum / area);
            }
        }

        context->Unmap(staging.Get(), 0);
        deskDupl->ReleaseFrame();

        int outIndex = 0;

        int top_strip_h = std::max(1, int(DST_H * 12 / 100));
        // top
        for (int i = 0; i < TOP_LEDS; ++i) {
            int sx0 = (DST_W * i) / TOP_LEDS;
            int sx1 = (DST_W * (i + 1)) / TOP_LEDS;
            uint64_t rsum=0, gsum=0, bsum=0;
            int cnt=0;
            for (int y=0;y<top_strip_h;++y) {
                for (int x=sx0;x<sx1;++x) {
                    int p=(y*DST_W + x)*3;
                    rsum += dst[p+0]; gsum += dst[p+1]; bsum += dst[p+2]; ++cnt;
                }
            }
            packet[outIndex++] = (uint8_t)(rsum / std::max(1,cnt));
            packet[outIndex++] = (uint8_t)(gsum / std::max(1,cnt));
            packet[outIndex++] = (uint8_t)(bsum / std::max(1,cnt));
        }

        // right
        for (int i = 0; i < RIGHT_LEDS; ++i) {
            int sy0 = (DST_H * i) / RIGHT_LEDS;
            int sy1 = (DST_H * (i + 1)) / RIGHT_LEDS;
            uint64_t rsum=0, gsum=0, bsum=0;
            int cnt=0;
            int x0 = std::max(0, DST_W - int(DST_W * 12 / 100));
            for (int y=sy0;y<sy1;++y) {
                for (int x=x0;x<DST_W;++x) {
                    int p=(y*DST_W + x)*3;
                    rsum += dst[p+0]; gsum += dst[p+1]; bsum += dst[p+2]; ++cnt;
                }
            }
            packet[outIndex++] = (uint8_t)(rsum / std::max(1,cnt));
            packet[outIndex++] = (uint8_t)(gsum / std::max(1,cnt));
            packet[outIndex++] = (uint8_t)(bsum / std::max(1,cnt));
        }

        // bottom (collect as triplets then append in reverse LED order)
        {
            std::vector<std::array<uint8_t,3>> bottombuf;
            bottombuf.reserve(BOTTOM_LEDS);
            int y0_start = std::max(0, DST_H - int(DST_H * 12 / 100));
            for (int i = 0; i < BOTTOM_LEDS; ++i) {
                int sx0 = (DST_W * i) / BOTTOM_LEDS;
                int sx1 = (DST_W * (i + 1)) / BOTTOM_LEDS;
                uint64_t rsum=0, gsum=0, bsum=0;
                int cnt=0;
                for (int y=y0_start;y<DST_H;++y) {
                    for (int x=sx0;x<sx1;++x) {
                        int p=(y*DST_W + x)*3;
                        rsum += dst[p+0]; gsum += dst[p+1]; bsum += dst[p+2]; ++cnt;
                    }
                }
                uint8_t rr = (uint8_t)(rsum / std::max(1,cnt));
                uint8_t gg = (uint8_t)(gsum / std::max(1,cnt));
                uint8_t bb = (uint8_t)(bsum / std::max(1,cnt));
                bottombuf.push_back({ rr, gg, bb });
            }
            for (int i = (int)bottombuf.size() - 1; i >= 0; --i) {
                packet[outIndex++] = bottombuf[i][0];
                packet[outIndex++] = bottombuf[i][1];
                packet[outIndex++] = bottombuf[i][2];
            }
        }

        // left (collect as triplets then append in reverse LED order)
        {
            std::vector<std::array<uint8_t,3>> leftbuf;
            leftbuf.reserve(LEFT_LEDS);
            int x1_end = std::min(DST_W, int(DST_W * 12 / 100));
            for (int i = 0; i < LEFT_LEDS; ++i) {
                int sy0 = (DST_H * i) / LEFT_LEDS;
                int sy1 = (DST_H * (i + 1)) / LEFT_LEDS;
                uint64_t rsum=0, gsum=0, bsum=0;
                int cnt=0;
                for (int y=sy0;y<sy1;++y) {
                    for (int x=0;x<x1_end;++x) {
                        int p=(y*DST_W + x)*3;
                        rsum += dst[p+0]; gsum += dst[p+1]; bsum += dst[p+2]; ++cnt;
                    }
                }
                uint8_t rr = (uint8_t)(rsum / std::max(1,cnt));
                uint8_t gg = (uint8_t)(gsum / std::max(1,cnt));
                uint8_t bb = (uint8_t)(bsum / std::max(1,cnt));
                leftbuf.push_back({ rr, gg, bb });
            }
            for (int i = (int)leftbuf.size() - 1; i >= 0; --i) {
                packet[outIndex++] = leftbuf[i][0];
                packet[outIndex++] = leftbuf[i][1];
                packet[outIndex++] = leftbuf[i][2];
            }
        }

        if (outIndex != NUM_LEDS * 3) {
            packet.assign(NUM_LEDS * 3, 0);
        }

        enhance_packet(packet);

        std::vector<uint8_t> sendbuf;
        if (cfg.use_header) sendbuf = make_framed(packet, frame_id);
        else sendbuf = packet;

        DWORD written = 0;
        if (hSerial != INVALID_HANDLE_VALUE) {
            if (!WriteFile(hSerial, sendbuf.data(), (DWORD)sendbuf.size(), &written, NULL)) {
                // serial write failed - maybe disconnected. Close and mark invalid.
                CloseHandle(hSerial);
                hSerial = INVALID_HANDLE_VALUE;
            } else {
                if (cfg.expect_ack) {
                    uint8_t ack = 0; DWORD rv = 0;
                    if (ReadFile(hSerial, &ack, 1, &rv, NULL) && rv == 1) {
                        // optional: do something on ack
                    } else {
                        // no ack
                    }
                }
            }
        } else {
            // attempt to reopen in background
            hSerial = openSerial(cfg.port, (DWORD)cfg.baud);
        }

        frame_id = (frame_id + 1) & 0xFF;

        std::this_thread::sleep_until(nextFrame);
        if (std::chrono::steady_clock::now() > nextFrame + std::chrono::seconds(1)) {
            nextFrame = std::chrono::steady_clock::now();
        }

        if (cfg.once) {
            // if once specified, send only one frame
            break;
        }
    } // while running

    if (hSerial != INVALID_HANDLE_VALUE) CloseHandle(hSerial);
    running = false;
}

// ------------------------------ start/stop helpers ------------------------------
void StartCaptureFromSettings() {
    std::scoped_lock lock(settings_mtx);
    if (running) return;
    should_stop = false;
    running = true;
    AppSettings snapshot = settings;
    captureThread = std::thread([snapshot]() {
        CaptureThreadFunc(snapshot);
    });
}

void StopCaptureAndJoin() {
    if (!running) return;
    should_stop = true;
    running = false;
    if (captureThread.joinable()) {
        captureThread.join();
    }
}

// ------------------------------ WinMain & message handling ------------------------------
LRESULT CALLBACK MainWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_CREATE:
        {
            // add tray icon
            NOTIFYICONDATA nid = {};
            nid.cbSize = sizeof(nid);
            nid.hWnd = hwnd;
            nid.uID = ID_TRAY_APP_ICON;
            nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
            nid.uCallbackMessage = WM_TRAYICON;
            nid.hIcon = LoadIcon(NULL, IDI_APPLICATION);
            wcscpy_s(nid.szTip, L"SyncLED Tray");
            Shell_NotifyIcon(NIM_ADD, &nid);
        }
        return 0;
    case WM_TRAYICON:
        if (LOWORD(lp) == WM_RBUTTONUP || LOWORD(lp) == WM_CONTEXTMENU) {
            ShowTrayMenu(hwnd);
        } else if (LOWORD(lp) == WM_LBUTTONDBLCLK) {
            // toggle start/stop on double-click
            if (running) {
                StopCaptureAndJoin();
            } else {
                StartCaptureFromSettings();
            }
        }
        return 0;
    case WM_COMMAND:
        switch (LOWORD(wp)) {
        case ID_TRAY_STARTSTOP:
            if (running) StopCaptureAndJoin();
            else StartCaptureFromSettings();
            break;
        case ID_TRAY_SETTINGS:
            CreateSettingsWindow(hwnd);
            break;
        case ID_TRAY_TOGGLE_HEADER:
            {
                std::scoped_lock lock(settings_mtx);
                settings.use_header = !settings.use_header;
                // don't auto restart for header change; it will apply next start
            }
            break;
        case ID_TRAY_TOGGLE_ACK:
            {
                std::scoped_lock lock(settings_mtx);
                settings.expect_ack = !settings.expect_ack;
            }
            break;
        case ID_TRAY_EXIT:
            StopCaptureAndJoin();
            // remove tray icon
            {
                NOTIFYICONDATA nid = {};
                nid.cbSize = sizeof(nid);
                nid.hWnd = hwnd;
                nid.uID = ID_TRAY_APP_ICON;
                Shell_NotifyIcon(NIM_DELETE, &nid);
            }
            PostQuitMessage(0);
            break;
        case 3001: // Apply from settings window
            if (settingsWnd && IsWindow(settingsWnd)) {
                std::scoped_lock lock(settings_mtx);
                settings.port = ReadStringFromEdit(editPort);
                settings.baud = ReadIntFromEdit(editBaud);
                settings.fps = ReadDoubleFromEdit(editFPS);
                settings.use_header = (SendMessage(chkHeader, BM_GETCHECK, 0, 0) == BST_CHECKED);
                settings.expect_ack = (SendMessage(chkAck, BM_GETCHECK, 0, 0) == BST_CHECKED);
                settings.test_glow_enable = (SendMessage(chkTestEnable, BM_GETCHECK, 0, 0) == BST_CHECKED);
                settings.test_r = ReadIntFromEdit(editR);
                settings.test_g = ReadIntFromEdit(editG);
                settings.test_b = ReadIntFromEdit(editB);
                // apply: if running, restart thread to pick up changes
                if (running) {
                    StopCaptureAndJoin();
                    StartCaptureFromSettings();
                }
            }
            break;
        case 3002: // Close settings
            DestroySettingsWindow();
            break;
        default:
            break;
        }
        return 0;
    case WM_DESTROY:
        {
            NOTIFYICONDATA nid = {};
            nid.cbSize = sizeof(nid);
            nid.hWnd = hwnd;
            nid.uID = ID_TRAY_APP_ICON;
            Shell_NotifyIcon(NIM_DELETE, &nid);
            PostQuitMessage(0);
        }
        return 0;
    default:
        return DefWindowProc(hwnd, msg, wp, lp);
    }
}

// ------------------------------ WinMain ------------------------------
int WINAPI WinMain(HINSTANCE hInst, HINSTANCE, LPSTR, int) {
    // initialize settings (optionally from a saved registry/file - omitted for brevity)
    // create hidden main window for tray messages
    WNDCLASS wc = {};
    wc.lpfnWndProc = MainWndProc;
    wc.hInstance = hInst;
    wc.lpszClassName = L"SyncLEDTrayClass";
    wc.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    RegisterClass(&wc);

    g_hwnd = CreateWindowEx(0, wc.lpszClassName, L"SyncLED Tray Window",
        WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 300, 200,
        NULL, NULL, hInst, NULL);
    if (!g_hwnd) return 1;

    // hide the main window since this is a background/tray app
    ShowWindow(g_hwnd, SW_HIDE);

    // message loop
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        // allow settings window child controls to process messages
        if (IsWindow(settingsWnd)) {
            // let controls handle as normal
        }
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    // cleanup
    StopCaptureAndJoin();
    DestroySettingsWindow();
    return 0;
}
