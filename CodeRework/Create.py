import tkinter as tk
from config import *

class Create(tk.Frame): #Initial Page Frame
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self['bg'] = BACKGROUND_COLOR

if __name__ == '__main__':
    exec(open("CodeRework/main.py").read())