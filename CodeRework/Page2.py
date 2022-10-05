import tkinter as tk

class Page2(tk.Frame): #Initial Page Frame
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self['bg'] = '#375280'