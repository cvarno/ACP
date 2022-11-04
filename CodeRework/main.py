import tkinter as tk

from config import *

from Home import Home
from Create import Create
from Modify import Modify
from Options import Options


class tkinterApp(tk.Tk): # Main Window Manager
    def __init__(self, *args, **kwargs):
    

        # __init__ function for class Tk
        tk.Tk.__init__(self, *args, **kwargs)
        
        # Create a frame to place other windows onto
        container = tk.Frame(self) 
        container.pack(side = "top", fill = "both", expand = True)
  
        container.grid_rowconfigure(0, weight = 1)
        container.grid_columnconfigure(0, weight = 1)

        # initializing frames to an empty array
        self.frames = {} 

        self.title(TITLE)
        self.geometry(SCREEN_SIZE_STRING)
        
        # iterating through a tuple consisting
        # of the different page layouts
        for F in (Home, Create, Modify, Options):
  
            # Create objects of each of the frames
            frame = F(container, self)
  
            # initializing frame of that object from startpage, page1, page2 respectively with for loop
            self.frames[F] = frame
            
            frame.grid(row = 0, column = 0, sticky ="nsew")
  

        # Show starting frame
        self.show_frame(Home)
  

        # Declare each frame so that the other frames can access them
        self.Home = Home
        self.Create = Create
        self.Modify = Modify
        self.Options = Options
        
    # to display the current frame passed as
    # parameter
    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()

app = tkinterApp()
app.mainloop()