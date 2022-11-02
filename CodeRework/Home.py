import tkinter as tk
from config import *

class Home(tk.Frame): #Initial Page Frame
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self['bg'] = BACKGROUND_COLOR
        message_L=tk.Label(self,  bg=BACKGROUND_COLOR,text='Welcome!\nWhat would you like to do?',fg='ghost white',font=("Verdana", 12),wraplength=0)
        message_L.grid(row = 0, column=0, columnspan = 100, padx = 5, pady = 5)

        go_create_B = tk.Button(self, text ="Create\nan Account",command = lambda : controller.show_frame(Create), width=10, height=5, font=("Verdana", 9))
        go_create_B.grid(row = 1, column = 0, padx = 5, pady = 5, sticky="NSEW")
        go_modify_B = tk.Button(self, text ="Modify/Disable\nan Account",command = lambda : controller.show_frame(Modify), width=10, height=5, font=("Verdana", 9))
        go_modify_B.grid(row = 2, column = 0, padx = 5, pady = 5, sticky="NSEW")
        go_settings_B = tk.Button(self, text ="Settings", command = lambda : controller.show_frame(Options), width=10, height=5, font=("Verdana", 9))
        go_settings_B.grid(row = 3, column = 0, padx = 5, pady = 5, sticky="NSEW")