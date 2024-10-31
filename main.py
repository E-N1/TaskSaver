from PyQt5 import QtWidgets
import subprocess
import json
import os
import sys
import ctypes
import psutil

def is_window_visible(window_handle):
    return ctypes.windll.user32.IsWindowVisible(window_handle)

def get_exe_path_from_window(window_handle):
    pid = ctypes.c_ulong()
    ctypes.windll.user32.GetWindowThreadProcessId(window_handle, ctypes.byref(pid))
    try:
        proc = psutil.Process(pid.value)
        return proc.exe()  # Get full executable path
    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
        return None

def get_open_applications():
    def enum_windows_proc(window_handle, _):
        if is_window_visible(window_handle):
            exe_path = get_exe_path_from_window(window_handle)
            if exe_path and exe_path not in app_tasks:
                app_tasks.append(exe_path)

    app_tasks = []
    wnd_enum_proc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_long, ctypes.c_long)(enum_windows_proc)
    ctypes.windll.user32.EnumWindows(wnd_enum_proc, 0)
    return app_tasks

class TaskManager(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()
        self.setWindowTitle("Task Manager")
        self.setGeometry(200, 200, 300, 100)

        # Define the path for the tasks.json file within the venv directory
        self.tasks_file = os.path.join(os.getcwd(), '.venv', 'tasks.json')

        # Ensure the venv directory exists
        os.makedirs(os.path.dirname(self.tasks_file), exist_ok=True)

        # Initialize the JSON file if it does not exist
        if not os.path.exists(self.tasks_file):
            with open(self.tasks_file, "w") as file:
                json.dump([], file)

    def initUI(self):
        self.layout = QtWidgets.QVBoxLayout()

        self.save_button = QtWidgets.QPushButton("Save")
        self.save_button.clicked.connect(self.save_task)
        self.layout.addWidget(self.save_button)

        self.open_button = QtWidgets.QPushButton("Open")
        self.open_button.clicked.connect(self.load_task)
        self.layout.addWidget(self.open_button)

        self.setLayout(self.layout)

    def save_task(self):
        task_name, ok = QtWidgets.QInputDialog.getText(self, 'Save Task', 'What should the task be called?')
        if ok and task_name:
            try:
                app_tasks = get_open_applications()

                task_info = {
                    "name": task_name,
                    "tasks": app_tasks
                }

                with open(self.tasks_file, "r") as file:
                    stored_tasks = json.load(file)

                stored_tasks.append(task_info)

                with open(self.tasks_file, "w") as file:
                    json.dump(stored_tasks, file, indent=4)

                QtWidgets.QMessageBox.information(self, "Success", f"Task '{task_name}' has been saved.")
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", f"Error saving task: {e}")

    def load_task(self):
        if not os.path.exists(self.tasks_file):
            QtWidgets.QMessageBox.warning(self, "Warning", "No tasks have been saved yet.")
            return

        with open(self.tasks_file, "r") as file:
            stored_tasks = json.load(file)

        task_names = [task["name"] for task in stored_tasks]

        task_name, ok = QtWidgets.QInputDialog.getItem(self, "Open Task", "Select a task:", task_names, 0, False)
        if ok and task_name:
            task_info = next((task for task in stored_tasks if task["name"] == task_name), None)
            if task_info:
                current_apps = get_open_applications()
                tasks_to_open = [task for task in task_info["tasks"] if task not in current_apps]

                for task in tasks_to_open:
                    try:
                        subprocess.Popen(task, shell=True)
                    except Exception as e:
                        QtWidgets.QMessageBox.warning(self, "Warning", f"Error opening: {task}\n{e}")

                QtWidgets.QMessageBox.information(self, "Tasks Loaded", f"Tasks have been opened.\n{', '.join(tasks_to_open)}")
            else:
                QtWidgets.QMessageBox.critical(self, "Error", "Task not found.")

def main():
    app = QtWidgets.QApplication(sys.argv)
    manager = TaskManager()
    manager.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()