import tkinter as tk

class Root(tk.Tk):
    def __init__(salf):
        super().__init__()

    salf.label=tk.Label(self,text='Hello world!')
    salf.label.pack

if __name__=='__main__':
    root=Root()
    root.mainloop()
