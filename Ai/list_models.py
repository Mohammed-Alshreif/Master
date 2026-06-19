#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
عرض الموديلات المتاحة من Gemini
"""
import sys
import io
import google.generativeai as genai

# Fix encoding for Windows
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

GEMINI_API_KEY = "AQ.Ab8RN6JKeGuB_r-UuA0Q0spsaj-guF7rEzh6yEpTby6BHO16nA"

def list_models():
    """عرض الموديلات المتاحة"""
    print("=" * 60)
    print("قائمة الموديلات المتاحة في Gemini")
    print("=" * 60 + "\n")
    
    try:
        genai.configure(api_key=GEMINI_API_KEY)
        
        for model in genai.list_models():
            print(f"[MODEL] {model.name}")
            print(f"        Version: {model.version}")
            print(f"        Description: {model.description}")
            if hasattr(model, 'supported_generation_methods'):
                print(f"        Methods: {model.supported_generation_methods}")
            print()
            
    except Exception as e:
        print(f"ERROR: {e}\n")

if __name__ == "__main__":
    list_models()
