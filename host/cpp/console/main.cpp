// syncled_fixed.cpp
// Build with Visual Studio: cl /EHsc /O2 syncled_fixed.cpp /link d3d11.lib dxgi.lib

#include <windows.h>
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

using Microsoft::WRL::ComPtr;

static const int DST_W = 128;
static const int DST_H = 128;
static const int NUM_LEDS = 96;
static const int TOP_LEDS = 31;
static const int RIGHT_LEDS = 17;
static const int BOTTOM_LEDS = 31;
static const int LEFT_LEDS = 17;

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

std::string nowStr() {
    using namespace std::chrono;
    auto t = system_clock::now();
    auto ms = duration_cast<milliseconds>(t.time_since_epoch()) % 1000;
    std::time_t tt = system_clock::to_time_t(t);
    char buf[64];
    tm local_tm;
    localtime_s(&local_tm, &tt);
    strftime(buf, sizeof(buf), "%Y-%m-%d %H:%M:%S", &local_tm);
    char out[96];
    sprintf_s(out, "%s.%03lld", buf, static_cast<long long>(ms.count()));
    return std::string(out);
}

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

int main(int argc, char* argv[]) {
    std::string port;
    int baud = 115200;
    double fps = 15.0;
    bool verbose = false;
    bool use_header = false;
    bool expect_ack = false;
    bool once = false;
    bool test_glow = false;
    int tr = -1, tg = -1, tb = -1;

    for (int i = 1; i < argc; ++i) {
        std::string a = argv[i];
        if (a == "--port" || a == "-p") { if (i+1<argc) port = argv[++i]; }
        else if (a == "--baud" || a == "-b") { if (i+1<argc) baud = atoi(argv[++i]); }
        else if (a == "--fps" || a == "-f") { if (i+1<argc) fps = atof(argv[++i]); }
        else if (a == "--verbose" || a == "-v") verbose = true;
        else if (a == "--header") use_header = true;
        else if (a == "--expect-ack") expect_ack = true;
        else if (a == "--once") once = true;
        else if (a == "--r" && i+1<argc) { tr = atoi(argv[++i]); test_glow = true; }
        else if (a == "--g" && i+1<argc) { tg = atoi(argv[++i]); test_glow = true; }
        else if (a == "--b" && i+1<argc) { tb = atoi(argv[++i]); test_glow = true; }
        else if (port.empty()) port = a;
    }

    if (port.empty()) {
        std::cerr << "Usage: syncled_fixed.exe --port COM3 [--baud 115200] [--fps 15] [--header] [--expect-ack] [--r RR --g GG --b BB] [--once] [--verbose]\n";
        return 1;
    }

    if (verbose) std::cout << nowStr() << " opening serial " << port << " @ " << baud << "\n";
    HANDLE hSerial = openSerial(port, (DWORD)baud);
    if (hSerial == INVALID_HANDLE_VALUE) {
        std::cerr << nowStr() << " Failed to open serial port " << port << "\n";
        return 1;
    }

    UINT createFlags = D3D11_CREATE_DEVICE_BGRA_SUPPORT;
    ComPtr<ID3D11Device> device;
    ComPtr<ID3D11DeviceContext> context;
    D3D_FEATURE_LEVEL fl;
    if (FAILED(D3D11CreateDevice(nullptr, D3D_DRIVER_TYPE_HARDWARE, nullptr, createFlags, nullptr, 0, D3D11_SDK_VERSION, &device, &fl, &context))) {
        std::cerr << nowStr() << " D3D11CreateDevice failed\n";
        CloseHandle(hSerial);
        return 1;
    }

    ComPtr<IDXGIDevice> dxgiDevice;
    if (FAILED(device.As(&dxgiDevice))) { CloseHandle(hSerial); return 1; }
    ComPtr<IDXGIAdapter> adapter;
    if (FAILED(dxgiDevice->GetAdapter(&adapter))) { CloseHandle(hSerial); return 1; }
    ComPtr<IDXGIOutput> output;
    if (FAILED(adapter->EnumOutputs(0, &output))) { CloseHandle(hSerial); return 1; }
    ComPtr<IDXGIOutput1> output1;
    if (FAILED(output.As(&output1))) { CloseHandle(hSerial); return 1; }
    ComPtr<IDXGIOutputDuplication> deskDupl;
    if (FAILED(output1->DuplicateOutput(device.Get(), &deskDupl))) {
        std::cerr << nowStr() << " DuplicateOutput failed\n";
        CloseHandle(hSerial);
        return 1;
    }

    ComPtr<ID3D11Texture2D> staging;
    int src_w = 0, src_h = 0;
    std::vector<int> xs(DST_W + 1), ys(DST_H + 1);
    std::vector<uint8_t> dst(DST_W * DST_H * 3);
    std::vector<uint8_t> packet;
    packet.resize(NUM_LEDS * 3);

    std::chrono::duration<double> frameIntervalD(1.0 / fps);
    auto frameInterval = std::chrono::duration_cast<std::chrono::steady_clock::duration>(frameIntervalD);
    auto nextFrame = std::chrono::steady_clock::now();

    if (verbose) std::cout << nowStr() << " capture started, target " << fps << " FPS\n";

    uint8_t frame_id = 0;

    if (test_glow) {
        int rr = (tr < 0) ? 128 : std::clamp(tr, 0, 255);
        int gg = (tg < 0) ? 128 : std::clamp(tg, 0, 255);
        int bb = (tb < 0) ? 128 : std::clamp(tb, 0, 255);
        std::vector<uint8_t> payload(NUM_LEDS * 3);
        for (int i = 0; i < NUM_LEDS; ++i) {
            payload[i*3 + 0] = (uint8_t)rr;
            payload[i*3 + 1] = (uint8_t)gg;
            payload[i*3 + 2] = (uint8_t)bb;
        }
        std::vector<uint8_t> outbuf;
        if (use_header) outbuf = make_framed(payload, frame_id);
        else outbuf = payload;
        DWORD written = 0;
        if (!WriteFile(hSerial, outbuf.data(), (DWORD)outbuf.size(), &written, NULL)) {
            std::cerr << nowStr() << " WriteFile failed for test glow\n";
        } else {
            if (verbose) std::cout << nowStr() << " test glow written bytes=" << written << "\n";
            if (expect_ack) {
                uint8_t ack = 0; DWORD rv = 0;
                if (ReadFile(hSerial, &ack, 1, &rv, NULL) && rv == 1) {
                    if (ack == 'A') std::cout << "ACK received\n";
                    else std::cout << "Received byte: " << std::hex << (int)ack << std::dec << "\n";
                } else {
                    if (verbose) std::cout << "No ACK\n";
                }
            }
        }
        frame_id = (frame_id + 1) & 0xFF;
        if (once) { CloseHandle(hSerial); return 0; }
    }

    while (true) {
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
            if (verbose) std::cout << nowStr() << " source resolution " << src_w << "x" << src_h << "\n";
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
            if (verbose) std::cerr << nowStr() << " warning: packet length mismatch " << outIndex << "\n";
            packet.assign(NUM_LEDS * 3, 0);
        }

        enhance_packet(packet);

        std::vector<uint8_t> sendbuf;
        if (use_header) sendbuf = make_framed(packet, frame_id);
        else sendbuf = packet;

        DWORD written = 0;
        if (!WriteFile(hSerial, sendbuf.data(), (DWORD)sendbuf.size(), &written, NULL)) {
            std::cerr << nowStr() << " WriteFile failed\n";
        } else {
            if (verbose) {
                static auto lastPrint = std::chrono::steady_clock::now();
                auto tnow = std::chrono::steady_clock::now();
                double ms = std::chrono::duration_cast<std::chrono::duration<double, std::milli>>(tnow - lastPrint).count();
                std::cout << nowStr() << " frame sent, bytes=" << written << " (frame dt ms ~ " << ms << ")\n";
                lastPrint = tnow;
            }
            if (expect_ack) {
                uint8_t ack = 0; DWORD rv = 0;
                if (ReadFile(hSerial, &ack, 1, &rv, NULL) && rv == 1) {
                    if (ack == 'A') {
                        if (verbose) std::cout << "ACK 'A' received\n";
                    } else {
                        if (verbose) std::cout << "Received byte: " << std::hex << (int)ack << std::dec << "\n";
                    }
                } else {
                    if (verbose) std::cout << "No ACK\n";
                }
            }
        }

        frame_id = (frame_id + 1) & 0xFF;

        std::this_thread::sleep_until(nextFrame);
        if (std::chrono::steady_clock::now() > nextFrame + std::chrono::seconds(1)) {
            nextFrame = std::chrono::steady_clock::now();
        }
    }

    CloseHandle(hSerial);
    return 0;
}
