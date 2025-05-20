import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

class Application:
    def __init__(self, test):
        self.root = tk.Tk()
        self.root.title("Data Standard")
        self.root.geometry("400x150")
        self.button = tk.Button(self.root, text="Select File", command = self.select_file)
        self.button.pack(pady=10)
        self.test = test
   
    def select_file(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.test.execute(file_path)
        else:
            return None
        
    def run(self):
        self.root.mainloop()