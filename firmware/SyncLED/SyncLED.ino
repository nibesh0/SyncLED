#include <FastLED.h>
#define NUM_LEDS 96
#define DATA_PIN 5
#define LED_BRIGHTNESS 255
CRGB leds[NUM_LEDS];
uint8_t payload[NUM_LEDS * 3];
enum State {H1, H2, FRAME, PAYLOAD, CHKS};
State st = H1;
uint8_t frame_id = 0;
int payload_index = 0;
uint8_t rx_frame_id = 0;
unsigned long last_byte_time = 0;
const unsigned long BYTE_TIMEOUT_MS = 200;

void setup() {
  Serial.begin(115200);
  FastLED.addLeds<WS2812B, DATA_PIN, GRB>(leds, NUM_LEDS);
  FastLED.setBrightness(LED_BRIGHTNESS);
  FastLED.show();
}   

void loop() {
  while (Serial.available()) {
    int b = Serial.read();
    if (b < 0) break;
    last_byte_time = millis();
    uint8_t ub = (uint8_t)b;
    if (st == H1) {
      if (ub == 0xAA) st = H2;
      else st = H1;
    } else if (st == H2) {
      if (ub == 0x55) st = FRAME;
      else st = H1;
    } else if (st == FRAME) {
      rx_frame_id = ub;
      payload_index = 0;
      st = PAYLOAD;
    } else if (st == PAYLOAD) {
      payload[payload_index++] = ub;
      if (payload_index >= NUM_LEDS * 3) st = CHKS;
    } else if (st == CHKS) {
      uint8_t chk = ub;
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
      st = H1;
    }
  }
  if (st != H1 && (millis() - last_byte_time) > BYTE_TIMEOUT_MS) {
    st = H1;
  }
}
