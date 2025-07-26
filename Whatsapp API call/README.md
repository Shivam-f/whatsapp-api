# Google Sheets to WhatsApp Screenshot Sender

This Python program automatically takes screenshots of Google Sheets and sends them to WhatsApp at scheduled intervals.

## Features

- ✅ Automatically opens Google Sheets in Chrome
- ✅ Takes screenshots of specific regions
- ✅ Opens WhatsApp Web for easy message sending
- ✅ Runs on a schedule (configurable intervals)
- ✅ Error handling and logging

## Setup Instructions

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Install Chrome Browser

Make sure you have Google Chrome installed on your system.

### 3. Configure the Script

Edit `app.py` and update these settings:

```python
# ================== USER CONFIGURATION ==================
google_sheet_url = "YOUR_GOOGLE_SHEET_URL_HERE"
screenshot_path = 'google_sheet_ss.png'
whatsapp_number = '+91XXXXXXXXXX'  # Your WhatsApp number
interval_minutes = 60  # How often to send (in minutes)
# ========================================================
```

### 4. Adjust Screenshot Region

In the `take_screenshot()` function, adjust the region coordinates:

```python
screenshot = pyautogui.screenshot(region=(100, 150, 1200, 700))
# Format: (x, y, width, height)
```

## Usage

### Run the Program

```bash
python app.py
```

### What Happens

1. The program opens your Google Sheet in Chrome
2. Takes a screenshot of the specified region
3. Opens WhatsApp Web with a pre-filled message
4. You manually attach the screenshot and send
5. Repeats every X minutes (as configured)

## Configuration Options

- **Google Sheet URL**: Your Google Sheets URL
- **WhatsApp Number**: Your phone number with country code
- **Screenshot Region**: Adjust coordinates to capture the right area
- **Interval**: How often to send screenshots (in minutes)

## Troubleshooting

### Screenshot Region

If the screenshot doesn't capture the right area:

1. Run the program
2. Note where your Google Sheet appears on screen
3. Adjust the region coordinates in `take_screenshot()`

### Chrome Issues

If Chrome doesn't open properly:

1. Make sure Chrome is installed
2. Check if Chrome is in your system PATH
3. Try updating Chrome to the latest version

### WhatsApp Issues

- Make sure your WhatsApp number is correct
- The program opens WhatsApp Web - you need to scan the QR code first time
- You'll need to manually attach the screenshot and send

## File Structure

```
├── app.py              # Main program
├── requirements.txt    # Dependencies
├── README.md          # This file
└── google_sheet_ss.png # Screenshots (created automatically)
```

## Notes

- The program runs continuously until stopped (Ctrl+C)
- Screenshots are saved as `google_sheet_ss.png`
- WhatsApp Web will open in your default browser
- You need to manually attach the screenshot and send the message
