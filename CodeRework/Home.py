import tkinter as tk
from config import *

class Home(tk.Frame): #Initial Page Frame
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self['bg'] = BACKGROUND_COLOR
        message_L=tk.Label(self,  bg=BACKGROUND_COLOR,text='Welcome!\nWhat would you like to do?',fg='ghost white',font=("Verdana", 12),wraplength=0)
        message_L.update()
        message_L.place(x=335, y=200)
        message_L.update()
        print(message_L.winfo_x(), message_L.winfo_y())

        go_create_B = tk.Button(self, text ="Create Account",command = lambda : controller.show_frame(controller.Create), width=15, height=2, font=("Verdana", 9))
        go_create_B.grid(column = 0, row = 0, padx = 108, pady = (200, 20))
        go_modify_B = tk.Button(self, text ="Change Existing",command = lambda : controller.show_frame(controller.Modify), width=15, height=2, font=("Verdana", 9))
        go_modify_B.grid(column = 0, row = 1, padx = 108, pady = 20)
        go_settings_B = tk.Button(self, text ="Settings", command = lambda : controller.show_frame(controller.Options), width=15, height=2, font=("Verdana", 9))
        go_settings_B.grid(column = 0, row = 2, padx = 108, pady = 20)


if __name__ == '__main__':
    exec(open("CodeRework/main.py").read())