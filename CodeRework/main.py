import tkinter as tk
from Home import Home
from Create import Create

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
        for F in (Home, Create):
  
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



app = tkinterApp()
app.mainloop()