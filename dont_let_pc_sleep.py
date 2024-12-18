import pyautogui
import time

# Define the interval in seconds (60 seconds = 1 minute)
interval = 60

while True:
    pyautogui.press('space')  # Simulate pressing the space bar
    print("Space bar pressed to prevent sleep.")
    time.sleep(interval)  # Wait for the specified interval before pressing again
