#include <FastLED.h>
#include <Wire.h>
#include <Adafruit_GFX.h>
#include <Adafruit_SSD1306.h>

#define NUM_LEDS 96
#define DATA_PIN 5
#define LED_BRIGHTNESS 255

CRGB leds[NUM_LEDS];
uint8_t payload[NUM_LEDS * 3];

enum State { H1, H2, FRAME, PAYLOAD, CHKS };
enum FrameType { FT_NONE, FT_LEDS, FT_STATUS };

State st = H1;
FrameType curFrameType = FT_NONE;

uint8_t frame_id = 0;
uint8_t rx_frame_id = 0;
int payload_index = 0;

char statusBuf[241];
int rx_status_len = 0;

unsigned long last_byte_time = 0;
const unsigned long BYTE_TIMEOUT_MS = 200;

#define SCREEN_WIDTH 128
#define SCREEN_HEIGHT 64
#define OLED_RESET -1
#define OLED_ADDR 0x3C

Adafruit_SSD1306 display(SCREEN_WIDTH, SCREEN_HEIGHT, &Wire, OLED_RESET);
bool oled_ok = false;

unsigned long last_oled_update = 0;

void drawBar(int x, int y, int w, int h, uint8_t pct) {
  display.drawRect(x, y, w, h, SSD1306_WHITE);
  int fw = (int)((w - 2) * (pct / 100.0));
  if (fw > 0) display.fillRect(x + 1, y + 1, fw, h - 2, SSD1306_WHITE);
}

void updateOLEDFromCSV(const char *s) {
  if (!oled_ok) return;
  char buf[241];
  strncpy(buf, s, 240);
  buf[240] = 0;
  char *save;
  char *t = strtok_r(buf, ",", &save);
  const char *dt = t ? t : "";
  t = strtok_r(NULL, ",", &save);
  int cpu = t ? atoi(t) : 0;
  t = strtok_r(NULL, ",", &save);
  int ram = t ? atoi(t) : 0;
  t = strtok_r(NULL, ",", &save);
  int gpu0 = t ? atoi(t) : 0;
  t = strtok_r(NULL, ",", &save);
  int gpu1 = t ? atoi(t) : 0;
  t = strtok_r(NULL, ",", &save);
  int dl = t ? atoi(t) : 0;
  t = strtok_r(NULL, ",", &save);
  int ul = t ? atoi(t) : 0;
  
  display.clearDisplay();
  
  // Row 0: CPU with bar
  display.setCursor(0, 0);
  display.print("CPU ");
  display.print(cpu);
  display.print("%");
  drawBar(50, 0, 78, 10, (uint8_t)constrain(cpu, 0, 100));
  
  // Row 1: RAM with bar  
  display.setCursor(0, 12);
  display.print("RAM ");
  display.print(ram);
  display.print("%");
  drawBar(50, 12, 78, 10, (uint8_t)constrain(ram, 0, 100));
  
  // Row 2: GPU0 and GPU1 with bars
  display.setCursor(0, 24);
  display.print("G0 ");
  display.print(gpu0);
  display.print("%");
  drawBar(30, 24, 30, 10, (uint8_t)constrain(gpu0, 0, 100));
  
  display.setCursor(64, 24);
  display.print("G1 ");
  display.print(gpu1);
  display.print("%");
  drawBar(94, 24, 34, 10, (uint8_t)constrain(gpu1, 0, 100));
  
  // Row 3: Network
  display.setCursor(0, 36);
  display.print("D ");
  display.print(dl*8);
  display.print("Mb/s U ");
  display.print(ul*8);
  display.print("Mb/s");
  
  // Row 4 (Last): Date/time
  display.setCursor(0, 54);
  display.print(dt);
  
  display.display();
}

void setup() {
  Serial.begin(115200);
  FastLED.addLeds<WS2812B, DATA_PIN, GRB>(leds, NUM_LEDS);
  FastLED.setBrightness(LED_BRIGHTNESS);
  FastLED.show();
  Wire.begin(21, 22);
  oled_ok = display.begin(SSD1306_SWITCHCAPVCC, OLED_ADDR);
  display.setRotation(2);
  display.clearDisplay();
  display.setTextSize(1);
  display.setTextColor(SSD1306_WHITE);
  display.setCursor(0, 0);
  display.println("LED SYNC READY");
  display.display();
}

void loop() {
  while (Serial.available()) {
    int b = Serial.read();
    if (b < 0) break;
    last_byte_time = millis();
    uint8_t ub = (uint8_t)b;
    if (st == H1) {
      if (ub == 0xAA) st = H2;
    } else if (st == H2) {
      if (ub == 0x55) { curFrameType = FT_LEDS; st = FRAME; }
      else if (ub == 0x56) { curFrameType = FT_STATUS; st = FRAME; }
      else { st = H1; curFrameType = FT_NONE; }
    } else if (st == FRAME) {
      if (curFrameType == FT_LEDS) {
        rx_frame_id = ub;
        payload_index = 0;
        st = PAYLOAD;
      } else if (curFrameType == FT_STATUS) {
        rx_status_len = ub;
        if (rx_status_len <= 0 || rx_status_len > 240) { st = H1; curFrameType = FT_NONE; }
        else { payload_index = 0; st = PAYLOAD; }
      } else {
        st = H1;
      }
    } else if (st == PAYLOAD) {
      if (curFrameType == FT_LEDS) {
        payload[payload_index++] = ub;
        if (payload_index >= NUM_LEDS * 3) st = CHKS;
      } else if (curFrameType == FT_STATUS) {
        statusBuf[payload_index++] = (char)ub;
        if (payload_index >= rx_status_len) { statusBuf[payload_index] = 0; st = CHKS; }
      } else st = H1;
    } else if (st == CHKS) {
      uint8_t chk = ub;
      if (curFrameType == FT_LEDS) {
        uint16_t s = rx_frame_id;
        for (int i = 0; i < NUM_LEDS * 3; ++i) s += payload[i];
        if (((uint8_t)s) == chk) {
          for (int i = 0; i < NUM_LEDS; ++i) {
            int j = i * 3;
            leds[i] = CRGB(payload[j], payload[j + 1], payload[j + 2]);
          }
          FastLED.show();
          Serial.write('A');
        } else {
          Serial.write('N');
        }
      } else if (curFrameType == FT_STATUS) {
        uint16_t s = 0x56;
        for (int i = 0; i < rx_status_len; ++i) s += (uint8_t)statusBuf[i];
        if (((uint8_t)s) == chk) {
          updateOLEDFromCSV(statusBuf);
          Serial.write('s');
        } else {
          Serial.write('n');
        }
      }
      st = H1;
      curFrameType = FT_NONE;
    }
  }
  if (st != H1 && (millis() - last_byte_time) > BYTE_TIMEOUT_MS) {
    st = H1;
    curFrameType = FT_NONE;
  }
}
