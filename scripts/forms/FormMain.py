# FormMain.py

import tkinter as tk

def launch():
    root = tk.Tk()
    root.title('Main Form')
    tk.Label(root, text='Hello!').pack()
    tk.Button(root, text='Click Me').pack()
    root.mainloop()
