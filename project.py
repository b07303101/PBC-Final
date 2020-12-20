# 期末專案

import tkinter as tk
import tkinter.font as tkFont

from PIL import ImageTk

class Meetings(tk.Frame):

    def __init__(self):
        tk.Frame.__init__(self)
        self.grid()
        self.createWidgets()
        
    def createWidgets(self):
        f1 = tkFont.Font(size = 30, family = "源泉圓體 M")
        f2 = tkFont.Font(size = 30, family = "源泉圓體 M")
        
        # self.lblTitle_A = tk.Label(self, text = "會議", height = 1, width = 10, font = f1)
        self.lblTitle_A = tk.Label(self, text = "會議", height = 1, width = 10, font = f1)
        self.btnCreate_New = tk.Button(self, text = "創建新會議", height = 1, width = 10, font = f2)
        
        self.lblTitle_A.grid(row = 0, column = 0, columnspan = 3, sticky = tk.NW)
        self.btnCreate_New.grid(row = 1, column = 2, sticky = tk.E)


call = Meetings()
call.master.title("開會小助手")  # 視窗名稱為開會小助手
call.master.geometry("1000x700")
call.mainloop()

