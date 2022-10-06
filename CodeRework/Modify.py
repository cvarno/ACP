import tkinter as tk

class Modify(tk.Frame): #Initial Page Frame
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self['bg'] = '#375280'

if __name__ == '__main__':
    exec(open("CodeRework/main.py").read())