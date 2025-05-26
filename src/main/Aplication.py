import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import threading
from Logger import Logger

class Application:
    def __init__(self, scenario):
        self.root = tk.Tk()
        self.root.title("Data Standard")
        self.root.geometry("400x150")

        self.button = tk.Button(self.root, text="Select File", command = self.select_file)
        self.button.pack(pady=5)

        self.labelProgress = tk.Label(self.root, text="")
        self.labelProgress.pack(pady=5)

        self.progress = ttk.Progressbar(self.root, orient='horizontal', length=200, mode='determinate')
        self.progress.pack(pady=5)

        self.labelTime = tk.Label(self.root, text="")
        self.labelTime.pack(pady=5)

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
        
    def run(self):
        self.root.mainloop()