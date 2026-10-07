"""
Local Server Runner for Hostel Housekeeping Attendance System
Starts a local web server and opens the portal in your default browser.
"""

import http.server
import socketserver
import webbrowser
import os
import sys

PORT = 8000
DIRECTORY = os.path.dirname(os.path.abspath(__file__))

class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

def start_server():
    os.chdir(DIRECTORY)
    with socketserver.TCPServer(("", PORT), Handler) as httpd:
        url = f"http://localhost:{PORT}"
        print(f"=======================================================")
        print(f" Hostel Housekeeping Attendance & Payroll System")
        print(f" Running locally at: {url}")
        print(f" Press Ctrl + C to stop the local server")
        print(f"=======================================================")
        try:
            webbrowser.open(url)
        except Exception:
            pass
        httpd.serve_forever()

if __name__ == "__main__":
    start_server()
