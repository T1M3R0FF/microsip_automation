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

        self.microsip_path = tk.StringVar()

        # Проверяем наличие файла конфигурации
        self.config_path = "config.ini"
        self.config = configparser.ConfigParser()

        if os.path.exists(self.config_path):
            self.config.read(self.config_path)
            saved_path = self.config.get("Settings", "MicroSIPPath", fallback="")
            self.microsip_path.set(saved_path)

        # Создаем Entry для ввода пути к MicroSIP
        entry_label = tk.Label(self.root, text="MicroSIP Path:")
        entry_label.grid(row=0, column=0, padx=10, pady=10, sticky="e")

        self.entry = tk.Entry(self.root, width=50, textvariable=self.microsip_path)
        self.entry.grid(row=0, column=1, padx=10, pady=10, ipadx=5, ipady=5)

        browse_button = tk.Button(self.root, text="Browse", command=self.browse_button)
        browse_button.grid(row=0, column=2, padx=10, pady=10, ipadx=10, ipady=5)

        # Создаем кнопку для указания пути
        set_path_button = tk.Button(self.root, text="Set Path", command=self.set_path)
        set_path_button.grid(row=1, column=1, pady=20, ipadx=20, ipady=10)

        # Создаем кнопки для старта и остановки
        button_frame = tk.Frame(self.root)
        button_frame.grid(row=2, column=1, pady=20)

        start_button = tk.Button(button_frame, text="Start", command=self.start)
        start_button.grid(row=0, column=0, padx=10, ipadx=20, ipady=10)

        stop_button = tk.Button(button_frame, text="Stop", command=self.stop)
        stop_button.grid(row=0, column=1, padx=10, ipadx=20, ipady=10)

        # Переменная для отслеживания статуса
        self.running = False

        self.root.mainloop()

    def browse_button(self):
        filename = filedialog.askopenfilename(title="Select MicroSIP Executable", filetypes=[("Executable files", "*.exe")])
        self.microsip_path.set(filename)

    def set_path(self):
        path = self.microsip_path.get()
        if not path:
            messagebox.showinfo("Info", "Please select the path to MicroSIP.")
        else:
            self.config["Settings"] = {"MicroSIPPath": path}
            with open(self.config_path, "w") as configfile:
                self.config.write(configfile)
            messagebox.showinfo("Info", f"MicroSIP Path set to: {path}")

    def open_application(self):
        try:
            process = subprocess.Popen([self.microsip_path.get()])
            process.wait()
        except Exception as e:
            messagebox.showerror("Error", f"Error opening MicroSIP: {str(e)}")

    def handle_hotkey(self):
        try:
            active_window = gw.getActiveWindow()
            active_window.activate()

            ctypes.windll.user32.keybd_event(0x11, 0, 0, 0)  # Нажимаем клавишу Ctrl
            ctypes.windll.user32.keybd_event(0x43, 0, 0, 0)  # Нажимаем клавишу C
            ctypes.windll.user32.keybd_event(0x43, 0, 2, 0)  # Отпускаем клавишу C
            ctypes.windll.user32.keybd_event(0x11, 0, 2, 0)  # Отпускаем клавишу Ctrl

            time.sleep(1)

            buf = pyperclip.paste()
            print(buf)

            self.open_application()

            kb.press_and_release('ctrl+a')
            kb.press_and_release('ctrl+v')
        except Exception as e:
            messagebox.showerror("Error", f"Error handling hotkey: {str(e)}")

    def start(self):
        if not self.running:
            kb.add_hotkey('ctrl+x', self.handle_hotkey)
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
