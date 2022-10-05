import tkinter as tk

class Home(tk.Frame): #Initial Page Frame
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self['bg'] = '#375280'