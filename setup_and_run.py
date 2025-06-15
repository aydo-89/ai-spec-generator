#!/usr/bin/env python3
"""
AI Spec Sheet Generator - Setup and Launch Script
This script will install dependencies and launch the app with ngrok
"""

import subprocess
import sys
import os

def run_cmd(cmd, description):
    """Run a command and handle errors"""
    print(f"🔧 {description}...")
    try:
        result = subprocess.run(cmd, shell=True, check=True, capture_output=True, text=True)
        print(f"✅ {description} completed")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ {description} failed: {e}")
        if e.stdout:
            print(f"Output: {e.stdout}")
        if e.stderr:
            print(f"Error: {e.stderr}")
        return False

def check_python_version():
    """Check if Python version is compatible"""
    version = sys.version_info
    if version.major < 3 or (version.major == 3 and version.minor < 8):
        print(f"❌ Python 3.8+ required. Current version: {version.major}.{version.minor}")
        return False
    print(f"✅ Python {version.major}.{version.minor} is compatible")
    return True

def main():
    print("🚀 AI Spec Sheet Generator - Setup & Launch")
    print("=" * 50)
    
    # Check Python version
    if not check_python_version():
        sys.exit(1)
    
    # Install dependencies
    print("\n📦 Installing dependencies...")
    
    dependencies = [
        "streamlit>=1.28.0",
        "pandas>=1.5.0", 
        "openpyxl>=3.0.0",
        "openai>=1.12.0",
        "pyngrok>=7.0.0"
    ]
    
    for dep in dependencies:
        if not run_cmd(f"pip3 install {dep}", f"Installing {dep.split('>=')[0]}"):
            print(f"⚠️  Failed to install {dep}")
    
    print("\n✅ Dependencies installed!")
    
    # Check if ngrok is available
    print("\n🔧 Checking ngrok...")
    ngrok_check = subprocess.run("which ngrok", shell=True, capture_output=True)
    if ngrok_check.returncode != 0:
        print("⚠️  ngrok not found. Installing via pip...")
        run_cmd("pip3 install pyngrok", "Installing pyngrok")
        print("💡 Note: You may want to install ngrok system-wide for better performance")
        print("   Visit: https://ngrok.com/download")
    else:
        print("✅ ngrok is available")
    
    print("\n" + "=" * 50)
    print("🎯 Setup complete! Ready to launch...")
    print("=" * 50)
    
    # Ask user if they want to launch
    response = input("\n🚀 Launch the AI Spec Sheet Generator now? (y/n): ").strip().lower()
    
    if response in ['y', 'yes']:
        print("\n🌟 Launching AI Spec Sheet Generator...")
        try:
            # Import and run the ngrok launcher
            exec(open('run_with_ngrok.py').read())
        except FileNotFoundError:
            print("❌ run_with_ngrok.py not found!")
            print("💡 Try running: python3 run_with_ngrok.py")
        except Exception as e:
            print(f"❌ Error launching: {e}")
    else:
        print("\n📋 To launch later, run: python3 run_with_ngrok.py")
        print("📋 Or locally: streamlit run app.py")

if __name__ == "__main__":
    main() 