import pyautogui
import keyboard
import time
import threading

# Global variable to track whether the script is paused or running
is_paused = False

def move_mouse():
    """Moves the mouse slightly every few seconds to prevent sleep."""
    while True:
        if not is_paused:
            # Move the mouse to a slightly different position to simulate activity
            pyautogui.move(10, 0)  # Move 10 pixels right
            pyautogui.move(-10, 0)  # Move 10 pixels left
            time.sleep(5)  # Wait for 5 seconds before moving again
        else:
            time.sleep(1)  # Wait a bit before checking again if paused

def toggle_pause():
    """Toggles the pause state when the spacebar is pressed."""
    global is_paused
    while True:
        if keyboard.is_pressed('space'):  # Check if spacebar is pressed
            is_paused = not is_paused  # Toggle the pause state
            if is_paused:
                print("Paused")
            else:
                print("Resumed")
            time.sleep(0.5)  # Debounce the key press

# Create threads to move the mouse and listen for the spacebar press
mouse_thread = threading.Thread(target=move_mouse)
pause_thread = threading.Thread(target=toggle_pause)

# Start the threads
mouse_thread.daemon = True
pause_thread.daemon = True
mouse_thread.start()
pause_thread.start()

# Keep the main thread alive
while True:
    time.sleep(1)
