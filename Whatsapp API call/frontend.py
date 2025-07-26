# --- Flask App for WhatsApp Screenshot Scheduler ---

# Import required libraries
from selenium.webdriver.common.keys import Keys
from flask import Flask, render_template_string, request, jsonify
import threading
import time
import schedule
from datetime import datetime, timedelta
import pyautogui
import webbrowser
import os
import pytz
import traceback
from PIL import Image
from selenium import webdriver
from selenium.webdriver.common.by import By
import chromedriver_autoinstaller
chromedriver_autoinstaller.install()

app = Flask(__name__)

# --- Spreadsheet Options ---
# Define the available Google Sheets and the region to capture for each
SPREADSHEETS = {
    'Sheet 1': {
        'url': "https://docs.google.com/spreadsheets/d/1_QSvPyOCCP43AZ6eZqNIm4OmJb8ds9EB4UD88C2-Sb4/edit#gid=909429816",
        'region': (200, 230, 1000, 500)
    },
    'Sheet 2': {
        'url': "https://docs.google.com/spreadsheets/d/1_QSvPyOCCP43AZ6eZqNIm4OmJb8ds9EB4UD88C2-Sb4/edit?gid=862699111#gid=862699111",
        'region': (100, 200, 1200, 600)  # Example region, adjust as needed
    }
}

# --- Shared State ---
# Used to keep track of the app's status and user selections
state = {
    'running': False,
    'thread': None,
    'last_status': 'Idle',
    'next_time': None,
    'whatsapp_number': '',
    'sheet_key': 'Sheet 1'  # Default selection
}

screenshot_path = "sheet_ss.png"

# --- Helper Functions ---


def get_number():
    # Return the WhatsApp number to send to, or a default
    return state['whatsapp_number'] or '+918651903000'


def get_sheet_url():
    # Return the URL of the selected Google Sheet
    return SPREADSHEETS[state['sheet_key']]['url']


def get_region():
    # Return the screenshot region for the selected sheet
    return SPREADSHEETS[state['sheet_key']]['region']


def zoom_out(times=4):
    import pyautogui
    import time
    for _ in range(times):
        pyautogui.hotkey('ctrl', '-')
        time.sleep(0.2)


# --- HTML for the Web Interface ---
# This is the user-facing web page
html = '''
<!DOCTYPE html>
<html>
<head>
    <title>WhatsApp Screenshot Scheduler</title>
    <style>
        body { font-family: Arial, sans-serif; max-width: 500px; margin: 50px auto; padding: 20px; }
        .container { background: #f5f5f5; padding: 30px; border-radius: 10px; }
        .timer { font-size: 24px; font-weight: bold; text-align: center; margin: 20px 0; }
        .btn { background: #007bff; color: white; padding: 15px 30px; border: none; border-radius: 5px; cursor: pointer; font-size: 16px; }
        .btn:disabled { background: #aaa; }
        .btn:hover:enabled { background: #0056b3; }
        .status { margin-top: 20px; text-align: center; }
        input[type=text] { padding: 10px; font-size: 16px; width: 100%; margin-bottom: 15px; }
        .dropdown { width: 100%; padding: 10px; font-size: 16px; margin-bottom: 15px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>📱 WhatsApp Screenshot Scheduler</h1>
        <form id="numberForm" onsubmit="event.preventDefault(); sendNow();">
            <label>WhatsApp Number (with country code):</label>
            <input type="text" id="whatsapp_number" name="whatsapp_number" placeholder="+91XXXXXXXXXX" required />
            <label for="sheet_key">Select Spreadsheet:</label>
            <select id="sheet_key" class="dropdown" name="sheet_key" onfocus="window.sheetKeyActive=true;" onblur="window.sheetKeyActive=false;">
                <option value="Sheet 1">Sheet 1</option>
                <option value="Sheet 2">Sheet 2</option>
            </select>
            <div style="text-align: center;">
                <button class="btn" id="sendBtn" type="submit">Send Now</button>
                <button class="btn" id="stopBtn" type="button" onclick="stopScheduler()">Stop</button>
            </div>
        </form>
        <div class="timer">
            <div id="timer">Next send: --:--</div>
        </div>
        <div class="status" id="status">Idle</div>
    </div>
    <script>
        // Send the screenshot now
        function sendNow() {
            const number = document.getElementById('whatsapp_number').value;
            const sheetKey = document.getElementById('sheet_key').value;
            fetch('/send_now', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ whatsapp_number: number, sheet_key: sheetKey })
            })
            .then(res => res.json())
            .then(data => {
                document.getElementById('status').textContent = data.status;
                updateStatus();
            });
        }
        // Stop the scheduler
        function stopScheduler() {
            fetch('/stop', {method: 'POST'})
                .then(res => res.json())
                .then(data => {
                    document.getElementById('status').textContent = data.status;
                    updateStatus();
                });
        }
        // Update the status and next send time
        function updateStatus() {
            fetch('/status')
                .then(res => res.json())
                .then(data => {
                    document.getElementById('timer').textContent = 'Next send: ' + data.next_time;
                    document.getElementById('status').textContent = data.status;
                    if (data.whatsapp_number) {
                        document.getElementById('whatsapp_number').value = data.whatsapp_number;
                    }
                    // Only update dropdown if user is not interacting with it
                    if (data.sheet_key && !window.sheetKeyActive) {
                        document.getElementById('sheet_key').value = data.sheet_key;
                    }
                });
        }
        setInterval(updateStatus, 5000);
        updateStatus();
    </script>
</body>
</html>
'''

# --- Core Logic Functions ---


def open_google_sheet():
    # Open the selected Google Sheet in the default browser
    webbrowser.open(get_sheet_url())
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Opening Google Sheet: {get_sheet_url()}")
    time.sleep(15)  # Wait for the sheet to load


def take_screenshot():
    # Take a screenshot of the specified region and save it using pyautogui
    region = get_region()
    import pyautogui
    import time
    import os
    zoom_out(times=5)  # Zoom out to ~50%
    time.sleep(1)      # Wait for zoom to finish
    screenshot = pyautogui.screenshot(region=region)
    screenshot.save(screenshot_path)
    # Ensure the file is fully written and accessible
    for _ in range(10):
        if os.path.exists(screenshot_path) and os.path.getsize(screenshot_path) > 0:
            try:
                with open(screenshot_path, 'rb') as f:
                    f.read(10)
                break
            except Exception:
                time.sleep(0.2)
        else:
            time.sleep(0.2)
    time.sleep(1)  # Extra delay to ensure file is ready
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Screenshot saved: {screenshot_path} (region: {region})")


def send_whatsapp_image():
    # Open the sheet, take a screenshot, and send it via WhatsApp using pywhatkit (with extra robustness)
    import time
    import pyautogui
    import pywhatkit
    import os
    try:
        open_google_sheet()
        take_screenshot()
        # Print working directory and screenshot path for debugging
        print(f"[send_whatsapp_image] CWD: {os.getcwd()}")
        abs_screenshot_path = os.path.abspath(screenshot_path)
        print(
            f"[send_whatsapp_image] Screenshot absolute path: {abs_screenshot_path}")
        if not os.path.exists(abs_screenshot_path):
            print(
                f"[send_whatsapp_image ERROR] Screenshot file does not exist: {abs_screenshot_path}")
            state['last_status'] = f'Error: Screenshot file not found: {abs_screenshot_path}'
            return
        # Bring browser window to front (best effort)
        try:
            pyautogui.hotkey('alt', 'tab')
            time.sleep(1)
        except Exception as e:
            print(
                f"[send_whatsapp_image] Could not bring browser to front: {e}")
        # Extra delay to ensure WhatsApp Web is loaded
        print("[send_whatsapp_image] Waiting 5 seconds for WhatsApp Web to load...")
        time.sleep(5)

        def try_send():
            send_func = getattr(pywhatkit, 'sendwhats_image', None)
            if send_func is None:
                send_func = getattr(pywhatkit, 'sendwhatsapp_image', None)
            if send_func is None:
                raise AttributeError(
                    'No WhatsApp image sending function found in pywhatkit')
            send_func(
                receiver=get_number(),
                img_path=abs_screenshot_path,
                caption="📊 Here is the latest data from the Sheet!",
                wait_time=60,  # Maximum wait for reliability
                tab_close=False
            )
        try:
            try_send()
            print(
                f"[{datetime.now().strftime('%H:%M:%S')}] Sent screenshot to WhatsApp via pywhatkit!")
            state['last_status'] = 'Screenshot sent!'
            # Wait for 5 seconds, then take a screenshot of the browser for debugging
            time.sleep(5)
            try:
                browser_ss = pyautogui.screenshot()
                browser_ss.save('browser_after_send.png')
                print(
                    '[send_whatsapp_image] Saved browser screenshot as browser_after_send.png')
            except Exception as e:
                print(
                    f"[send_whatsapp_image] Could not take browser screenshot: {e}")
        except Exception as e:
            print(
                f"[send_whatsapp_image RETRY] First send failed, retrying in 3 seconds: {e}")
            time.sleep(3)
            try_send()
            print(
                f"[{datetime.now().strftime('%H:%M:%S')}] Sent screenshot to WhatsApp via pywhatkit! (after retry)")
            state['last_status'] = 'Screenshot sent after retry!'
            time.sleep(5)
            try:
                browser_ss = pyautogui.screenshot()
                browser_ss.save('browser_after_send.png')
                print(
                    '[send_whatsapp_image] Saved browser screenshot as browser_after_send.png (after retry)')
            except Exception as e:
                print(
                    f"[send_whatsapp_image] Could not take browser screenshot after retry: {e}")
        # Wait and redirect to localhost
        time.sleep(2)
        import webbrowser
        webbrowser.open('http://127.0.0.1:5050')
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        print(f"[send_whatsapp_image OUTER ERROR] {tb}")
        state['last_status'] = f'send_whatsapp_image outer error: {e}\n{tb}'

# --- Flask Routes ---


@app.route('/')
def index():
    # Render the main web page
    return render_template_string(html)


@app.route('/send_now', methods=['POST'])
def send_now():
    # Handle the "Send Now" button: update state and start sending in a thread
    data = request.get_json()
    number = data.get('whatsapp_number', '').strip()
    sheet_key = data.get('sheet_key', 'Sheet 1')
    if number:
        state['whatsapp_number'] = number
    if sheet_key in SPREADSHEETS:
        state['sheet_key'] = sheet_key
    state['last_status'] = 'Sending screenshot...'
    threading.Thread(target=send_whatsapp_image, daemon=True).start()
    # Start the schedule if not already running
    if not state['running']:
        start_schedule()
    return jsonify({'status': state['last_status']})


@app.route('/stop', methods=['POST'])
def stop():
    # Stop the scheduler
    stop_schedule()
    return jsonify({'status': state['last_status']})


@app.route('/status')
def status():
    # Return the current status and next scheduled send time
    if state['next_time']:
        next_time_str = state['next_time'].strftime('%Y-%m-%d %H:%M:%S IST')
    else:
        next_time_str = '--:--'
    return jsonify({'status': state['last_status'], 'next_time': next_time_str, 'whatsapp_number': state['whatsapp_number'], 'sheet_key': state['sheet_key']})

# --- Scheduling Functions ---


def schedule_ist(hour, minute):
    # Schedule a job at a specific hour and minute (IST)
    ist = pytz.timezone('Asia/Kolkata')

    def job():
        now = datetime.now(ist)
        print(
            f"[Scheduler] Triggered at {now.strftime('%Y-%m-%d %H:%M:%S %Z')}")
        state['last_status'] = f'Sending screenshot at {now.strftime('%H:%M %Z')}'
        send_whatsapp_image()
        update_next_time()
    schedule.every().day.at(f"{hour:02d}:{minute:02d}").do(job)


def update_next_time():
    # Update the next scheduled send time
    ist = pytz.timezone('Asia/Kolkata')
    now = datetime.now(ist)
    times = []
    for h, m in [(11, 0), (17, 0)]:
        t = now.replace(hour=h, minute=m, second=0, microsecond=0)
        if t < now:
            t += timedelta(days=1)
        times.append(t)
    next_time = min(times)
    state['next_time'] = next_time


def scheduler_loop():
    # Run the scheduler loop in a background thread
    while state['running']:
        schedule.run_pending()
        time.sleep(1)


def start_schedule():
    # Start the daily schedule for sending screenshots
    schedule.clear()
    schedule_ist(11, 0)
    schedule_ist(17, 0)
    update_next_time()
    state['running'] = True
    t = threading.Thread(target=scheduler_loop, daemon=True)
    state['thread'] = t
    t.start()
    state['last_status'] = 'Scheduler running (11:00, 17:00 IST)'


def stop_schedule():
    # Stop the scheduler
    state['running'] = False
    schedule.clear()
    state['next_time'] = None
    state['last_status'] = 'Stopped'


# --- App Entry Point ---
if __name__ == '__main__':
    print("✅ Frontend running on port 5050")
    import threading as _threading

    def _open_browser():
        time.sleep(2)
        webbrowser.open('http://localhost:5050')
    _threading.Thread(target=_open_browser, daemon=True).start()
    app.run(debug=False, port=5050)
