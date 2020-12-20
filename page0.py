import tkinter as tk

class Meetings(tk.Frame):
    def __init__(self):
        tk.Frame.__init__(self)
        self.grid()
      
    
        
        
call = Meetings()
call.master.title("開會小助手")  
call.master.geometry('1000x700')
caption = tk.Label(call, text='開會小助手', bg='yellow', font=('Arial', 12), width=30, height=2)
caption.pack()
get_start = tk.Button(call, text='開始使用', font=('Arial', 12), width=10, height=1)
get_start.pack()
call.mainloop()