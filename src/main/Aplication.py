from enum import Enum
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import threading
from Logger import Logger
from Define import *

class Application:
    class MesageType(Enum):
        ERROR = 1
        WARNING = 2
        INFO = 3

    def __init__(self, scenario):
        self.root = tk.Tk()
        self.root.title(APP_NAME)
        self.root.geometry(APP_SIZE)

        self.button = tk.Button(self.root, text=BUTTON_TEXT, command = self.select_file)
        self.button.pack(pady=PADDING)

        self.labelProgress = tk.Label(self.root, text="")
        self.labelProgress.pack(pady=PADDING)

        self.progress = ttk.Progressbar(self.root, orient='horizontal', length=PROGRESSBAR_LENGTH, mode='determinate')
        self.progress.pack(pady=PADDING)

        self.labelTime = tk.Label(self.root, text="")
        self.labelTime.pack(pady=PADDING)

        self.scenario = scenario
   
    def select_file(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            thread = threading.Thread(target=self.executeScenario, args=(file_path,))
            thread.start()
        else:
            return None
        
    def executeScenario(self, file_path):
        self.scenario.execute(file_path, self)

    def setProgress(self, current, total):
        ratio = (current / total) * 100
        self.progress['value'] = ratio
        self.labelProgress.config(text=f"Progress: {current}/{total} ({ratio:.2f}%)")

    def setExecuteTime(self, time):
        self.labelTime.config(text=f"Execution Time: {str(time).split('.')[0]}")

    def showMessagebox(self, messageType,  title, message):
        if messageType == self.MesageType.INFO:
            messagebox.showinfo(title, message)
        elif messageType == self.MesageType.WARNING:
            messagebox.showwarning(title, message)
        elif messageType == self.MesageType.ERROR:
            messagebox.showerror(title, message)
        
    def run(self):
        self.root.mainloop()