import os
from datetime import datetime
import comtypes
import cv2
import time
import numpy as np
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.chrome.options import Options
import winsound
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from mss import mss
import threading
from ctypes import POINTER, cast
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioMeterInformation
from selenium.webdriver.common.by import By
import openpyxl
from openpyxl.styles import Font, Alignment

# --------------------- SETTINGS ---------------------
# === CONFIG ===
STREAM_TESTER_URL = "https://www.radiantmediaplayer.com/stream-tester.html"
STREAMING_URL = "https://cdn01khi.tamashaweb.com:8087/YlUHeDQb7a/city41News-abr/playlist.m3u8"
SCREEN_REGION = {"top": 200, "left": 200, "right": 200, "width": 1920, "height": 1200}  # Adjust to your stream area
ALERT_SOUND = "alert.mp3"
EXCEL_LOG_FILE = f"freeze_log_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
EMAIL_ALERT = True
EMAIL_FROM = "expressdartt@gmail.com"
EMAIL_TO = "alibasti2021@gmail.com"
EMAIL_CC = "m.minds456@gmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_USERNAME = "expressdartt@gmail.com"
SMTP_PASSWORD = "ztonyoyowxsdsarc"


# ----------------------------------------------------


def play_alert():
    try:
        print("[AUDIO] Playing beep alert...")
        winsound.Beep(1000, 500)  # Frequency: 1000 Hz, Duration: 500ms
    except Exception as e:
        print(f"[AUDIO] Beep failed: {e}")


def send_email(subject, body, attachment_path=None):
    if not EMAIL_ALERT:
        return
    try:
        msg = MIMEMultipart()
        msg["From"] = EMAIL_FROM
        msg["To"] = EMAIL_TO
        msg["CC"] = EMAIL_CC
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        if attachment_path:
            with open(attachment_path, "rb") as f:
                from email.mime.base import MIMEBase
                from email import encoders

                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename={attachment_path}",
                )
                msg.attach(part)

        # ✅ Fix: Send to both TO and CC recipients
        recipients = [EMAIL_TO] + EMAIL_CC.split(",")  # split CC addresses by comma
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USERNAME, SMTP_PASSWORD)
            server.sendmail(EMAIL_FROM, recipients, msg.as_string())  # send to "TO + CC" list
        print("[EMAIL] Alert sent with attachment." if attachment_path else "[EMAIL] Alert sent.")
    except Exception as e:
        print(f"[EMAIL] Failed to send alert: {e}")


def initialize_excel_log():
    if not os.path.exists(EXCEL_LOG_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Freeze Events"
        ws.append(["Timestamp", "Event Description"])
        header_font = Font(bold=True)
        for cell in ws[1]:
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
        wb.save(EXCEL_LOG_FILE)
        print(f"[EXCEL] Created new log file: {EXCEL_LOG_FILE}")
    else:
        print(f"[EXCEL] Logging to existing file: {EXCEL_LOG_FILE}")


def setup_browser():
    options = Options()
    options.add_argument("--start-maximized")
    # options.add_argument("--headless=new")  # ✅ Chrome headless with screen rendering
    driver = webdriver.Chrome(options=options)
    driver.get(STREAM_TESTER_URL)
    print("[INFO] Browser launched and loading stream...")
    print(driver.page_source)
    # === Scroll and Paste the Streaming URL ===
    try:
        # Replace placeholder text below with actual element info
        input_field = driver.find_element(By.XPATH, "//*[@id='stream-url']")
        driver.execute_script("arguments[0].scrollIntoView(true);", input_field)
        time.sleep(1)
        input_field.clear()
        input_field.send_keys(STREAMING_URL)
        input_field.send_keys(Keys.ENTER)  # Optional: hit Enter to start
        print("[INFO] Streaming URL entered.")

        play_streaming = driver.find_element(By.XPATH, "//*[@id='rmp']/button")
        play_streaming.click()
        print("[INFO] Play Button Clicked")

    except Exception as e:
        print(f"[ERROR] Failed to enter stream URL: {e}")
    time.sleep(8)  # Allow stream to load
    return driver


def is_stream_frozen(prev_frame, curr_frame, threshold=500):
    diff = cv2.absdiff(prev_frame, curr_frame)
    gray = cv2.cvtColor(diff, cv2.COLOR_BGR2GRAY)
    non_zero_count = cv2.countNonZero(gray)
    return non_zero_count < threshold


# Replace the log_freeze_event function
def log_freeze_event(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")

    wb = openpyxl.load_workbook(EXCEL_LOG_FILE)
    ws = wb.active
    ws.append([timestamp, message])
    wb.save(EXCEL_LOG_FILE)


# Monitor audio for mute (low volume)
def monitor_audio():
    print("[AUDIO] Monitoring audio started...")
    comtypes.CoInitialize()
    try:
        threshold = 0.02
        silence_duration = 2
        silent_start = None
        alerted = False

        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioMeterInformation._iid_, CLSCTX_ALL, None)  # This is correct
        meter = cast(interface, POINTER(IAudioMeterInformation))

        while True:
            volume = meter.GetPeakValue()

            if volume < threshold:
                if silent_start is None:
                    silent_start = time.time()
                elif time.time() - silent_start >= silence_duration and not alerted:
                    send_email("Audio Muted Alert", "Audio has been muted for more than 2 seconds.")
                    alerted = True
            else:
                silent_start = None
                alerted = False

            time.sleep(0.1)
    finally:
        comtypes.CoUninitialize()


def monitor_stream(region):
    sct = mss()
    prev_frame = None
    freeze_count = 0
    max_freeze_tolerance = 1
    print("[INFO] Monitoring stream...")

    while True:
        img = sct.grab(region)
        full_res_frame = np.array(img)
        full_res_frame = cv2.cvtColor(full_res_frame, cv2.COLOR_BGRA2BGR)

        # Resize for faster processing
        frame = cv2.resize(full_res_frame, (1200, 720))

        if prev_frame is not None:
            if is_stream_frozen(prev_frame, frame):
                freeze_count += 1
                log_freeze_event(f"Freeze detected...")
                if freeze_count >= max_freeze_tolerance:
                    log_freeze_event("[ALERT] Stream appears to be frozen!")

                    # Save full-res screenshot
                    freeze_image_path = "freeze_screenshot.jpg"
                    cv2.imwrite(freeze_image_path, full_res_frame, [int(cv2.IMWRITE_JPEG_QUALITY), 80])

                    # Play alert sound in separate thread
                    threading.Thread(target=play_alert, daemon=True).start()

                    # Send email with screenshot
                    send_email(
                        subject="Stream Freeze Detected",
                        body="The monitored stream appears to be frozen. See attached screenshot.",
                        attachment_path=freeze_image_path
                    )
                    freeze_count = 0
            else:
                freeze_count = 0

        prev_frame = frame
        time.sleep(0.5)


def main():
    initialize_excel_log()
    driver = setup_browser()

    audio_thread = threading.Thread(target=monitor_audio, daemon=True)
    audio_thread.start()

    monitor_thread = threading.Thread(target=monitor_stream, args=(SCREEN_REGION,), daemon=True)
    monitor_thread.start()

    print("[INFO] Stream and audio monitoring started. Press Ctrl+C to stop.")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("[INFO] Stopping monitoring... Exiting.")
        driver.quit()


if __name__ == "__main__":
    main()
