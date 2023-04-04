import pyautogui
import random
import time
import msvcrt
import matplotlib.pyplot as plt
import keyboard

# Define the variables for counting and timing
count = 0
total_time = 0

# Create lists to store the x and y coordinates of the mouse movements
x_coords = []
y_coords = []

# Display a message to let you know that the script is running
print("The script is now running. Press Ctrl-C to stop.")

while True:
    # Generate random x and y coordinates for the cursor to move to
    x = random.randint(0, pyautogui.size().width)
    y = random.randint(0, pyautogui.size().height)

    # Move the mouse cursor to the new position
    pyautogui.moveTo(x, y, duration=0.25)

    # Append the x and y coordinates to the lists
    x_coords.append(x)
    y_coords.append(y)

    # Increment the count and total time
    count += 1
    total_time += 0.25

    # Convert the total time to minutes and seconds
    total_time_minutes = int(total_time / 60)
    total_time_seconds = int(total_time % 60)

    # Display how many times and how much time it has moved the cursor
    print(
        f"The cursor has been moved {count} times for a total of {total_time_minutes} minutes and {total_time_seconds} seconds.")

    # Ask the user if they want to stop the script
    print("Press 's' to stop the script.")
    start_time = time.monotonic()
    while True:
        # Wait for 10 seconds for a response
        if time.monotonic() - start_time > 5:
            break

        # Check if there's any input available to be read from stdin
        if msvcrt.kbhit():
            # Read the input and break out of the loop
            if msvcrt.getch().decode('utf-8').lower() == "s":
                break
            else:
                # Move the cursor back to the center of the screen
                pyautogui.moveTo(pyautogui.size().width / 2, pyautogui.size().height / 2, duration=0.25)

    # Wait for 30 seconds before the next movement
    time.sleep(10)

    # Break out of the loop if the user pressed 's'
    if keyboard.is_pressed('s'):
        break

# Plot the x and y coordinates on a scatter plot
plt.scatter(x_coords, y_coords, s=5, c='blue')
plt.title('Mouse Movement Map')
plt.xlabel('X Coordinate')
plt.ylabel('Y Coordinate')
plt.show()
