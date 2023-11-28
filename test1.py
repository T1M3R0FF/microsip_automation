import time
import pygetwindow as gw
import pyperclip
import ctypes
import subprocess
import keyboard as kb
import tkinter as tk
from tkinter import filedialog, messagebox
import configparser
import os


class MicroSIPAutomation:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("MicroSIP Automation")

        self.microsip_paths = [tk.StringVar() for _ in range(5)]

        # Check for the existence of the configuration file
        self.config_path = "config.ini"
        self.config = configparser.ConfigParser()

        if os.path.exists(self.config_path):
            self.config.read(self.config_path)
            for i in range(5):
                saved_path = self.config.get("Settings", f"MicroSIPPath{i+1}", fallback="")
                self.microsip_paths[i].set(saved_path)

        # Create Entry widgets for inputting paths
        for i in range(5):
            entry_label = tk.Label(self.root, text=f"MicroSIP Path {i+1}:")
            entry_label.grid(row=i, column=0, padx=10, pady=10, sticky="e")

            entry = tk.Entry(self.root, width=50, textvariable=self.microsip_paths[i])
            entry.grid(row=i, column=1, padx=10, pady=10, ipadx=5, ipady=5)

            browse_button = tk.Button(self.root, text=f"Browse {i+1}", command=lambda idx=i: self.browse_button(idx))
            browse_button.grid(row=i, column=2, padx=10, pady=10, ipadx=10, ipady=5)

        # Create Set Path buttons for each path
        for i in range(5):
            set_path_button = tk.Button(self.root, text=f"Set Path {i+1}", command=lambda idx=i: self.set_path(idx))
            set_path_button.grid(row=i + 5, column=1, pady=10, ipadx=20, ipady=10)

        # Create buttons for start and stop
        button_frame = tk.Frame(self.root)
        button_frame.grid(row=10, column=1, pady=20)

        start_button = tk.Button(button_frame, text="Start", command=self.start)
        start_button.grid(row=0, column=0, padx=10, ipadx=20, ipady=10)

        stop_button = tk.Button(button_frame, text="Stop", command=self.stop)
        stop_button.grid(row=0, column=1, padx=10, ipadx=20, ipady=10)

        # Variable to track the status
        self.running = False

        self.root.mainloop()

    def browse_button(self, idx):
        filename = filedialog.askopenfilename(title=f"Select MicroSIP {idx+1} Executable", filetypes=[("Executable files", "*.exe")])
        self.microsip_paths[idx].set(filename)

    def set_path(self, idx):
        path = self.microsip_paths[idx].get()
        if not path:
            messagebox.showinfo("Info", f"Please select the path to MicroSIP {idx+1}.")
        else:
            self.config["Settings"][f"MicroSIPPath{idx+1}"] = path
            with open(self.config_path, "w") as configfile:
                self.config.write(configfile)
            messagebox.showinfo("Info", f"MicroSIP Path {idx+1} set to: {path}")

    def open_application(self, idx):
        try:
            process = subprocess.Popen([self.microsip_paths[idx].get()])
            process.wait()
        except Exception as e:
            messagebox.showerror("Error", f"Error opening MicroSIP {idx+1}: {str(e)}")

    def handle_hotkey(self, idx):
        try:
            active_window = gw.getActiveWindow()
            active_window.activate()

            ctypes.windll.user32.keybd_event(0x11, 0, 0, 0)  # Press Ctrl
            ctypes.windll.user32.keybd_event(0x43, 0, 0, 0)  # Press C
            ctypes.windll.user32.keybd_event(0x43, 0, 2, 0)  # Release C
            ctypes.windll.user32.keybd_event(0x11, 0, 2, 0)  # Release Ctrl

            time.sleep(1)

            buf = pyperclip.paste()

            if self.microsip_paths[idx].get():
                self.open_application(idx)
                kb.press_and_release('ctrl+a')
                kb.press_and_release('ctrl+v')
            else:
                messagebox.showinfo("Info", f"Path {idx + 1} is not set. Please set the path.")
        except Exception as e:
            messagebox.showerror("Error", f"Error handling hotkey {idx + 1}: {str(e)}")

    def start(self):
        if not self.running:
            hotkeys = ['Z', 'X', 'C', 'V', 'B']

            for i in range(5):
                kb.add_hotkey(f'ctrl+{hotkeys[i]}', lambda idx=i: self.handle_hotkey(idx))

            self.running = True
            try:
                time.sleep(0.01)
                messagebox.showinfo("Info", "MicroSIP Automation started.")
            except Exception as e:
                messagebox.showerror("Error", f"Error starting MicroSIP Automation: {str(e)}")

    def stop(self):
        if self.running:
            try:
                kb.unhook_all()
                self.running = False
                messagebox.showinfo("Info", "MicroSIP Automation stopped.")
            except Exception as e:
                messagebox.showerror("Error", f"Error stopping MicroSIP Automation: {str(e)}")


if __name__ == "__main__":
    app = MicroSIPAutomation()
