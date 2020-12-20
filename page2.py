import tkinter as tk

class Meetings(tk.Frame):
    def __init__(self):
        tk.Frame.__init__(self)
        self.grid()
        self.createWidgets()

    def createWidgets(self):
        self.lblname = tk.Label(self, text = "會議名稱", height = 1, width = 10)
        self.txtname = tk.Text(self, height = 1, width = 20)
        self.lblname.grid(column=350, row=500, sticky=tk.N + tk.E + tk.W + tk.S)
        self.txtname.grid(column=351, row=500, sticky=tk.N + tk.E + tk.W + tk.S)
     
    
call = Meetings()
call.master.title("開會小助手")  
call.master.geometry('1000x700')
call.mainloop()