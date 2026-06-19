#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
اختبار Gemini API
"""
import sys
import io
import google.generativeai as genai
import json

# Fix encoding for Windows
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

GEMINI_API_KEY = "AQ.Ab8RN6JKeGuB_r-UuA0Q0spsaj-guF7rEzh6yEpTby6BHO16nA"

def test_gemini_simple():
    """اختبار بسيط مع Gemini"""
    print("=" * 60)
    print("اختبار Gemini API")
    print("=" * 60)
    
    try:
        # تفعيل الـ API
        genai.configure(api_key=GEMINI_API_KEY)
        model = genai.GenerativeModel("gemini-2.5-flash")
        
        print("\n[OK] Connected to Gemini successfully!\n")
        
        # السؤال الأول - بسيط
        print("-" * 60)
        print("Question 1: Calculate average humidity values")
        print("-" * 60)
        
        response1 = model.generate_content(
            "Calculate the average of these values: 95, 92, 88, 90. What is the result?"
        )
        print(f"Gemini Response:\n{response1.text}\n")
        
        # السؤال الثاني - قرار للنظام
        print("-" * 60)
        print("Question 2: Irrigation system decision")
        print("-" * 60)
        
        response2 = model.generate_content(
            """You are an expert in automated irrigation systems.
            Current sensor data:
            - Temperature: 28.8°C
            - Humidity: 95%
            - Soil Moisture: 100%
            - Water Level: 22 cm
            
            Give a short recommendation (2 sentences) and a yes/no decision to turn on the motor.
            Respond ONLY with valid JSON (no markdown):
            {
              "recommendation": "...",
              "motor_on": true/false
            }"""
        )
        print(f"Gemini Response:\n{response2.text}\n")
        
        # محاولة parse JSON
        try:
            text = response2.text.strip()
            # تنظيف الـ markdown
            if "```json" in text:
                text = text.split("```json")[1].split("```")[0]
            elif "```" in text:
                text = text.split("```")[1].split("```")[0]
            
            data = json.loads(text)
            print("[OK] JSON parsed successfully:")
            print(f"    - Recommendation: {data.get('recommendation', 'N/A')}")
            print(f"    - Turn on motor: {data.get('motor_on', 'N/A')}\n")
        except:
            print("[WARNING] Could not parse JSON\n")
        
        # السؤال الثالث - اختبار الأوامر
        print("-" * 60)
        print("Question 3: Device control commands")
        print("-" * 60)
        
        response3 = model.generate_content(
            """Give me JSON commands to control these devices:
            - Servo device (on/off)
            - Motor (on/off)
            - Motor Speed (0-100)
            
            Response format (ONLY JSON, no markdown):
            {
              "actions": {
                "Servo device": true,
                "Motor": false,
                "Motor Speed": 75
              }
            }"""
        )
        print(f"Gemini Response:\n{response3.text}\n")
        
        print("=" * 60)
        print("[OK] Test completed successfully!")
        print("=" * 60)
        
    except Exception as e:
        print(f"\n[ERROR] {e}\n")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    test_gemini_simple()
