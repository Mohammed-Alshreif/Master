#include <Arduino.h>
#include <WiFi.h>
#include <Firebase_ESP_Client.h>
#include <DHT.h>
#include <math.h>
#include <esp_task_wdt.h>

//=========================================
#define UPdateRate 10
#define READ_INTERVAL_MS 5000
unsigned long lastReadMillis = 0;
unsigned long SystemCounter = 0;
#define irrigationTimeon 1
#define irrigationWAITTime 5  //WAIT*READ_INTERVAL_MS
#define FertilizeTime 1       //WAIT*READ_INTERVAL_MS
//=========================================
// GPIO pins
#define SOIL_PIN 34
#define DHT_PIN 4
#define DHT_TYPE DHT11
#define TRIG_PIN 5
#define ECHO_PIN 18
#define IRRIGATION_RELAY_PIN 19
#define FERTILIZER_RELAY_PIN 21
#define LED_PIN 2
//==========================================

// Helper functions
#include "addons/TokenHelper.h"
#include "addons/RTDBHelper.h"

// Wi-Fi credentials
#define WIFI_SSID "WE_54CFA8"
#define WIFI_PASSWORD "d0299d76"

// Firebase credentials
#define API_KEY "AIzaSyDEslUsBv8uOyXIXqsbehPDqgHx2PCUYIY"
#define DATABASE_URL "https://esptest1-edcd8-default-rtdb.firebaseio.com/"
//===========================================
//===========================================
FirebaseJson json;
long waterLevel = 0;
int soil = 0;
float temp = 0;
float hum = 0;
bool IrrigationState = false;
bool FertilizerState = false;

int Irrigationmin_Hum = 20;
int Irrigation_maxHum = 80;

//ldr flux sensor
const int LDR_PIN = 35;
long adc = 0;
float lux = 0;
//=====================
// Firebase
FirebaseData fbdo;
FirebaseAuth auth;
FirebaseConfig config;


bool signupOK = false;

DHT dht(DHT_PIN, DHT_TYPE);

// fixed paths
String basePath = "/apartments/flat1/rooms/";
//===========================================

typedef enum IrrigationStat {
  IRR_PUMP_ON,
  IRR_WAIT_SOIL,
  IRR_FINISHED
} IrrigationStatmachine_TD;
IrrigationStatmachine_TD IrrigationStatmachine = IRR_PUMP_ON;
//=========================================================
//=========================================================
bool reconnectWiFi();
bool reconnectFirebase();
void updateConnections();
void updateSensors();
void irrigationTask();
void FertilizeTask();
//=========================================================
long readUltrasonicFiltered() {
  const int samples = 7;
  long values[samples];


  for (int i = 0; i < samples; i++) {
    digitalWrite(TRIG_PIN, LOW);
    delayMicroseconds(2);
    digitalWrite(TRIG_PIN, HIGH);
    delayMicroseconds(10);
    digitalWrite(TRIG_PIN, LOW);

    long duration = pulseIn(ECHO_PIN, HIGH, 30000);  // max 30ms
    values[i] = duration * 0.034 / 2;                // يتحول سم
    delay(20);
  }

  // Bubble sort
  for (int i = 0; i < samples - 1; i++) {
    for (int j = i + 1; j < samples; j++) {
      if (values[i] > values[j]) {
        long tmp = values[i];
        values[i] = values[j];
        values[j] = tmp;
      }
    }
  }


  long medianValue = values[samples / 2];
  long sum = 0;
  int count = 0;
  for (int i = 0; i < samples; i++) {
    if (values[i] > medianValue * 0.8 && values[i] < medianValue * 1.2) {
      sum += values[i];
      count++;
    }
  }

  if (count > 0) {
    return sum / count;  // متوسط القيم المعقولة
  } else {
    return medianValue;  // fallback
  }
}


//========================= UPDATE SENSORS ================
void updateSensors() {
  // ====== Read sensors ======
  soil = 100 - analogRead(SOIL_PIN) / 40.95;
  temp = dht.readTemperature();
  hum = dht.readHumidity();
  //waterLevel = readUltrasonicFiltered();
  waterLevel =20; //(waterLevel > 100) ? 100 : waterLevel;
  adc = 0;
  for (int i = 0; i < 100; i++) {
    adc += analogRead(LDR_PIN);
  }
  adc = adc / 100;
  lux = 25 * exp(adc * 0.00192);
  //lux=lux/1000;
  // ====== Upload to Firebase ======
  // // Living room sensors
  // Firebase.RTDB.setInt(&fbdo, basePath + "room1/sensors/sensors1/value", temp);
  // Firebase.RTDB.setInt(&fbdo, basePath + "room1/sensors/sensors2/value", hum);

  // // Room1 soil + water +lux
  // Firebase.RTDB.setInt(&fbdo, basePath + "room1/sensors/sensors3/value", soil);
  // Firebase.RTDB.setString(&fbdo, basePath + "room1/sensors/sensors3/name", "soil");
  // Firebase.RTDB.setString(&fbdo, basePath + "room1/sensors/sensors3/unit", "%");

  // Firebase.RTDB.setInt(&fbdo, basePath + "room1/sensors/sensors4/value", waterLevel);
  // Firebase.RTDB.setString(&fbdo, basePath + "room1/sensors/sensors4/name", "waterLevel");
  // Firebase.RTDB.setString(&fbdo, basePath + "room1/sensors/sensors4/unit", "cm");

  // Firebase.RTDB.setInt(&fbdo, basePath + "room1/sensors/sensors5/value",(long) lux);
  // Firebase.RTDB.setString(&fbdo, basePath + "room1/sensors/sensors5/name", "lux sensor");
  // Firebase.RTDB.setString(&fbdo, basePath + "room1/sensors/sensors5/unit", "lux");
  json.clear();

  json.set("sensors/sensors1/value", temp);
  json.set("sensors/sensors1/name", "Temperature");
  json.set("sensors/sensors1/unit", "°C");

  json.set("sensors/sensors2/value", hum);
  json.set("sensors/sensors2/name", "Air Humidity");
  json.set("sensors/sensors2/unit", "%");

  json.set("sensors/sensors3/value", soil);
  json.set("sensors/sensors3/name", "Soil Moisture");
  json.set("sensors/sensors3/unit", "%");

  json.set("sensors/sensors4/value", waterLevel);
  json.set("sensors/sensors4/name", "WaterLevel");
  json.set("sensors/sensors4/unit", "cm");

  json.set("sensors/sensors5/value", round((lux / 1000.0) * 1000) / 1000.0);
  json.set("sensors/sensors5/name", "Lux sensor");
  json.set("sensors/sensors5/unit", "10^3 lux");

  json.set("sensors/sensors6/value", (short)Irrigationmin_Hum);
  json.set("sensors/sensors6/name", "Irrigation_MIN_Hum AI STATE");
  json.set("sensors/sensors6/unit", "%");
  
  json.set("sensors/sensors7/value", (short)Irrigation_maxHum);
  json.set("sensors/sensors7/name", "Irrigation MAX_Hum AI STATE");
  json.set("sensors/sensors7/unit", "%");
  // ====== Single Firebase Request ======
  if (!Firebase.RTDB.updateNode(
        &fbdo,
        basePath + "room1",
        &json)) {
    Serial.printf(
      "UpdateNode Error: %s\n",
      fbdo.errorReason().c_str());
    ESP.restart();
  }
  Serial.printf("Soil=%d Temp=%.1f Hum=%.1f Water=%ld  lux=%.1f  adc= %d\n", soil, temp, hum, waterLevel, lux, adc);
}

//==============================================================
//==========================Read_FireBase=======================
void Read_FireBase() {
   FirebaseJsonData result;

    if (Firebase.RTDB.getJSON(&fbdo, basePath + "room1"))
    {
        FirebaseJson &json = fbdo.jsonObject();

        json.get(result, "devices/device3/status");
        IrrigationState = result.boolValue;

        json.get(result, "devices/device3/Max_HUM");
        Irrigation_maxHum = result.intValue;

        json.get(result, "devices/device3/MIN_HUM");
        Irrigationmin_Hum = result.intValue;

        json.get(result, "devices/device1/status");
        FertilizerState = result.boolValue;

        Serial.printf(
            "IrrigationState=%d Irrigation_maxHum=%d Irrigationmin_Hum=%d FertilizerState=%d\n",
            IrrigationState,
            Irrigation_maxHum,
            Irrigationmin_Hum,
            FertilizerState);
    }
    else
    {
        Serial.println("Failed to read JSON");
        Serial.println(fbdo.errorReason());
        ESP.restart();
    }
}

//====================================================================
//=============================FertilizeTask==========================
void FertilizeTask() {
  static int FertilizeCounter = 0;
  if (FertilizerState == true) {
    FertilizeCounter++;
    if (FertilizeCounter <= FertilizeTime) {
      digitalWrite(FERTILIZER_RELAY_PIN, LOW);  //FERTILIZER_RELAY_PIN on
      Serial.println("FERTILIZER Pump ON ---");
    } else {
      digitalWrite(FERTILIZER_RELAY_PIN, HIGH);  //FERTILIZER_RELAY_PIN on
      Serial.println("FERTILIZER Pump OFF ---");
      FertilizeCounter = 0;
      FertilizerState = false;
      Firebase.RTDB.setBool(&fbdo, basePath + "room1/devices/device1/status", FertilizerState);
    }

  } else {
    FertilizeCounter = 0;
  }
}
//====================================================================
//====================================================================
void irrigationTask() {
  static int irrigationCounter = 0;
  //===================== Irrigation State Machine ===================
  if (IrrigationState == true) {
    irrigationCounter++;
    Read_FireBase();
    switch (IrrigationStatmachine) {
      case IRR_PUMP_ON:

        if (irrigationCounter <= irrigationTimeon) {
          digitalWrite(IRRIGATION_RELAY_PIN, LOW);  //IRRIGATION_RELAY_PIN on
          Serial.println("Pump ON ---");
        } else {
          IrrigationStatmachine = IRR_WAIT_SOIL;
          irrigationCounter = 0;
          Serial.println("Pump OFF ---");
          digitalWrite(IRRIGATION_RELAY_PIN, HIGH);  //IRRIGATION_RELAY_PIN on
        }

        break;

      case IRR_WAIT_SOIL:

        if (irrigationCounter >= irrigationWAITTime) {
          irrigationCounter = 0;
          IrrigationStatmachine = IRR_PUMP_ON;
          Serial.printf("updateSensors \n");
          updateSensors();
          Serial.printf(
            "IRR_WAIT_SOIL\n --Soil after irrigation = %d\n",
            soil);

          if (soil >= Irrigation_maxHum) {
            IrrigationState = false;
            Firebase.RTDB.setBool(&fbdo, basePath + "room1/devices/device3/status", IrrigationState);
            Serial.println(
              "Target humidity reached");
          } else {

            Serial.println(
              "Need more water");
          }
        }

        break;

      default:
        break;
    }
  } else {
    irrigationCounter = 0;
    Serial.println("Pump OFF ---");
    digitalWrite(IRRIGATION_RELAY_PIN, HIGH);  //IRRIGATION_RELAY_PIN oFF
  }
}
//=========================================================
//==============================================================SETUP==============================================================
void setup() {
  Serial.begin(9600);
  pinMode(TRIG_PIN, OUTPUT);
  pinMode(IRRIGATION_RELAY_PIN, OUTPUT);
  digitalWrite(IRRIGATION_RELAY_PIN, HIGH);
  pinMode(FERTILIZER_RELAY_PIN, OUTPUT);
  digitalWrite(FERTILIZER_RELAY_PIN, HIGH);
  pinMode(LED_PIN, OUTPUT);
  digitalWrite(LED_PIN, HIGH);
  pinMode(ECHO_PIN, INPUT);
  analogReadResolution(12);  // Set ADC resolution to 12-bit (0 - 4095)
  dht.begin();
  // connect Wi-Fi
  WiFi.mode(WIFI_STA);
  WiFi.setTxPower(WIFI_POWER_19_5dBm); 
  WiFi.begin(WIFI_SSID, WIFI_PASSWORD);
  Serial.print("Connecting to Wi-Fi");
  while (WiFi.status() != WL_CONNECTED) {
    Serial.print(".");
    delay(300);
  }
  Serial.println();
  Serial.println("Connected with IP: " + WiFi.localIP().toString());

  // Firebase
  config.api_key = API_KEY;
  config.database_url = DATABASE_URL;
  if (Firebase.signUp(&config, &auth, "", "")) {
    signupOK = true;
    Serial.println("Firebase sign-up OK");
  } else {
    Serial.printf("Firebase sign-up failed: %s\n", config.signer.signupError.message.c_str());
  }
  config.token_status_callback = tokenStatusCallback;
  Firebase.begin(&config, &auth);
  Firebase.reconnectWiFi(true);
  digitalWrite(LED_PIN, LOW);
  //watch dog timer
  //  esp_task_wdt_init(30, true);
  //  esp_task_wdt_add(NULL);
}
//==============================================================
//==============================================================
//==============================================================LOOP==============================================================
void loop() {
  //esp_task_wdt_reset();
  if (Firebase.ready() && signupOK && (millis() - lastReadMillis > READ_INTERVAL_MS)) {
    lastReadMillis = millis();
    SystemCounter++;
    updateConnections();
    if (SystemCounter > UPdateRate) {
      SystemCounter=0;
      Read_FireBase();
      updateSensors();
    }

    irrigationTask();
    FertilizeTask();
    Serial.printf("Free heap: %d\n", ESP.getFreeHeap());
    ToggleLed();
    Serial.print("RSSI: ");
    Serial.print(WiFi.RSSI());
    Serial.println(" dBm");
  }else{
    //ESP.restart();
  }
}
//==============================================================
//==============================================================
void updateConnections() {
  // if (!connectWiFi()) {
  //   reconnectWiFi();
  //}
  // if (!connectFirebase()) {
  //   reconnectFirebase();
  // }
}  //==============================================================
//===============================================================
bool reconnectWiFi() {
  if (WiFi.status() == WL_CONNECTED)
    return true;

  Serial.println("[WIFI] Reconnecting...");

  WiFi.disconnect();
  WiFi.begin(WIFI_SSID, WIFI_PASSWORD);

  unsigned long startTime = millis();

  while (WiFi.status() != WL_CONNECTED && millis() - startTime < 15000) {
    delay(250);
  }

  return WiFi.status() == WL_CONNECTED;
}
//==============================================================
//==============================================================
bool reconnectFirebase() {
  if (!Firebase.ready()) {
    Firebase.begin(&config, &auth);
    Firebase.reconnectWiFi(true);
  }

  return Firebase.ready();
}
//==============================================================
//==============================================================
void ToggleLed()
{
    static bool ledState = false;

    ledState = !ledState;
    digitalWrite(LED_PIN, ledState);
}