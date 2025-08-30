#main.py
import tkinter as tk
from controller import Controller   



if __name__ == "__main__":
    root = tk.Tk()
    app_controller = Controller(root)
    root.mainloop()