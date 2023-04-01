import pyautogui
import time
import math

def move_mouse_in_circle(radius, duration=1, steps=100):
    center_x, center_y = pyautogui.position()
    angle_step = 2 * math.pi / steps

    start_time = time.time()
    for i in range(steps + 1):
        angle = i * angle_step
        x = center_x + radius * math.cos(angle)
        y = center_y + radius * math.sin(angle)
        pyautogui.moveTo(x, y)
        current_time = time.time()
        elapsed_time = current_time - start_time
        if elapsed_time < duration:
            time.sleep((duration - elapsed_time) / steps)


while True:
    move_mouse_in_circle(radius=100, duration=2, steps=100)
    time.sleep(30)
