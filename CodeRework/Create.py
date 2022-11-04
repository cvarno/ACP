import tkinter as tk
from config import *

class Create(tk.Frame): #Initial Page Frame
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self['bg'] = BACKGROUND_COLOR

        go_home_B = tk.Button(self, text ="Home",command = lambda : controller.show_frame(controller.Home))
        go_home_B.grid(row = 0, column = 0, padx = 5, pady = 5, sticky="NW")

if __name__ == '__main__':
    exec(open("CodeRework/main.py").read())