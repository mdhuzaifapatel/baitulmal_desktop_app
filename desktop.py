import webview
import threading
import uvicorn
import os
import sys

from app import app

def start_server():
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="error")

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

if __name__ == "__main__":
    try:
        # Start FastAPI in a background thread
        threading.Thread(target=start_server, daemon=True).start()

        # Create the application window
        # Note: icon parameter is often picky on Windows, so we omit it here 
        # and rely on the PyInstaller icon flag for the taskbar/window.
        webview.create_window(
            "Hazrat Sher-E-Sawar (RH) Baitulmal Trust®",
            "http://127.0.0.1:8000",
            width=1400,
            height=900,
            maximized=True
        )
        
        # Start the webview using the reliable Edge (Chromium) engine
        webview.start(gui='edge')
    except Exception as e:
        import traceback
        with open("crash_log.txt", "w") as f:
            f.write(traceback.format_exc())