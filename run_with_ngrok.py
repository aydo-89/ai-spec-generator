#!/usr/bin/env python3
import subprocess
import threading
import time
import os
from pyngrok import ngrok, conf
import signal
import sys

def signal_handler(sig, frame):
    """Handle Ctrl+C gracefully"""
    print('\n🛑 Shutting down...')
    ngrok.kill()
    sys.exit(0)

def run_streamlit():
    """Run the Streamlit app"""
    try:
        cmd = [sys.executable, "-m", "streamlit", "run", "app.py", "--server.port=8501"]
        subprocess.run(cmd, check=True)
    except subprocess.CalledProcessError as e:
        print(f"❌ Error running Streamlit: {e}")
    except KeyboardInterrupt:
        pass

def main():
    """Main function to set up ngrok tunnel and run Streamlit"""
    signal.signal(signal.SIGINT, signal_handler)
    
    print("🚀 Starting AI Spec Sheet Generator with ngrok...")
    print("=" * 60)
    
    # Kill any existing ngrok processes
    try:
        ngrok.kill()
    except:
        pass
    
    # Set up ngrok
    try:
        # You can set ngrok auth token here if needed
        # conf.get_default().auth_token = "your_ngrok_auth_token"
        
        # Create tunnel
        print("🔗 Creating ngrok tunnel...")
        http_tunnel = ngrok.connect(8501)
        print(f"✅ Tunnel created: {http_tunnel.public_url}")
        print(f"📱 Share this URL: {http_tunnel.public_url}")
        print("=" * 60)
        
        # Start Streamlit in a separate thread
        streamlit_thread = threading.Thread(target=run_streamlit)
        streamlit_thread.daemon = True
        streamlit_thread.start()
        
        print("🤖 Streamlit app is starting...")
        print("⏱️  Give it a moment to initialize...")
        time.sleep(3)
        
        print(f"\n🌐 Your AI Spec Sheet Generator is now live at:")
        print(f"🔗 {http_tunnel.public_url}")
        print("\n📋 Instructions:")
        print("1. Click the link above to access your app")
        print("2. Upload your three files (Dev Log, Template, BOM)")
        print("3. Click 'Generate Spec Sheets' to process")
        print("4. Download your results!")
        print("\n⏹️  Press Ctrl+C to stop the server")
        print("=" * 60)
        
        # Keep the main thread alive
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            pass
            
    except Exception as e:
        print(f"❌ Error setting up ngrok: {e}")
        print("💡 Make sure you have ngrok installed and configured")
        return
    
    finally:
        print("\n🛑 Shutting down ngrok...")
        ngrok.kill()

if __name__ == "__main__":
    main() 