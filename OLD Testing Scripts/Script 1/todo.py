# Define an empty list to store tasks
tasks = []

# Define a function to add tasks to the list
def add_task(task):
    tasks.append(task)
    print("Task added: ", task)

# Define a function to remove tasks from the list
def remove_task(task):
    if task in tasks:
        tasks.remove(task)
        print("Task removed: ", task)
    else:
        print("Task not found: ", task)

# Define a function to print the current list of tasks
def print_tasks():
    print("Current tasks:")
    for task in tasks:
        print("- ", task)

# Main loop to prompt user for input
while True:
    print("\nWhat would you like to do?")
    print("1. Add a task")
    print("2. Remove a task")
    print("3. Print tasks")
    print("4. Quit")

    choice = input("> ")
    if choice == "1":
        task = input("Enter a new task: ")
        add_task(task)
    elif choice == "2":
        task = input("Enter the task to remove: ")
        remove_task(task)
    elif choice == "3":
        print_tasks()
    elif choice == "4":
        break
    else:
        print("Invalid choice. Please enter 1-4.")
