import tkinter as tk
from PIL import Image, ImageTk

class Home(tk.Frame): #Initial Page Frame
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self['bg'] = '#375280'
        message_L=tk.Label(self,  bg='#375280',text='Welcome!\nWhat would you like to do?',fg='ghost white',font=("Verdana", 12),wraplength=0)
        message_L.grid(row = 0, column=0, columnspan = 100, padx = 5, pady = 5)

if __name__ == '__main__':
    exec(open("CodeRework/main.py").read())