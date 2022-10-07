import tkinter as tk

from config import *

from Create import Create
from Modify import Modify
from Options import Options

class tkinterApp(tk.Tk): #Driver Code

    # __init__ function for class tkinterApp
    def __init__(self, *args, **kwargs):
    

        # __init__ function for class Tk
        tk.Tk.__init__(self, *args, **kwargs)
        
        # creating a container
        container = tk.Frame(self) 
        container.pack(side = "top", fill = "both", expand = True)
  
        container.grid_rowconfigure(0, weight = 1)
        container.grid_columnconfigure(0, weight = 1)

        # initializing frames to an empty array
        self.frames = {} 

        self.title('YMCA Account Creation Program')
        self.geometry('900x600')
        
        # iterating through a tuple consisting
        # of the different page layouts
        for F in (Home, Create, Modify, Options):
  
            frame = F(container, self)
  
            # initializing frame of that object from startpage, page1, page2 respectively with for loop
            self.frames[F] = frame
            
            frame.grid(row = 0, column = 0, sticky ="nsew")
  
        self.show_frame(Home)
  
    # to display the current frame passed as
    # parameter
    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()

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

app = tkinterApp()
app.mainloop()