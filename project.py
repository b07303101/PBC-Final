# # 期末專案

import tkinter as tk
# 期末專案

import tkinter as tk
import tkinter.font as tkFont

from PIL import ImageTk

root = tk.Tk()

class Page0:
    def __init__(self, master=None):
        self.root = master
        self.page0 = tk.Frame(self.root, width = 1000, height = 700)
        self.page0.configure(bg = 'Dimgray')
        self.page0.master.title("開會小助手")
        self.page0.grid()
        self.createPage_0()

    def createPage_0(self):
        f1 = tkFont.Font(size = 45, family = "源泉圓體 B")
        f2 = tkFont.Font(size = 20, family = "源泉圓體 M")
        color_1 = 'Steelblue'
        color_2 = 'Dimgray'
        self.lblCaption = tk.Label(self.page0, text= '開 會 小 助 手', bg = color_2, fg = 'white', font = f1, width = 30, height  = 2, anchor = 'center')
        self.btnStart = tk.Button(self.page0, text = '開始使用', bg = 'Goldenrod', font = f2, width = 20, height = 1, command = self.click_btnStart)
        
        # self.lblCaption.place(x = 140, y = 200)
        self.lblCaption.place(relx = 0.5, rely = 0.5, anchor = 'center')
        self.btnStart.place(relx = 0.5, y = 480, anchor = 'center')
   
    def click_btnStart(self):
        self.page0.destroy()
        PageA()

class PageA:

    def __init__(self, master = None):
        self.root = master
        self.pageA = tk.Frame(self.root, width = 1000, height = 700)
        self.pageA.master.title("會議")
        self.pageA.grid()
        self.createPage_A()
        
    def createPage_A(self):
        f1 = tkFont.Font(size = 30, family = "源泉圓體 B")
        f2 = tkFont.Font(size = 20, family = "源泉圓體 M")
        # color_1 = 'Steelblue'
        # color_2 = 'Dimgray'
        self.lblTitle_A = tk.Label(self.pageA, text = " 會議", height = 1 , width = 15, font = f1, bg = 'DarkSlateGray', fg = 'white', anchor = 'w')
        self.btnCreate_New = tk.Button(self.pageA, text = "創建新會議", height = 1, width = 10, font = f2, bg = 'Snow', fg = 'black', command = self.click_btnCreate_New)
        
        self.lblTitle_A.place(x = 0, y = 50)
        self.btnCreate_New.place(x = 800, y = 550)
    
    
    def newMeeting(self):
        self.name = meeting_name.get()
        self.btnCreate_New = tk.Button(self.pageA, text = self.name, height = 1, width = 10, font = f2)
        self.btnCreate_New.place(x = 30, y = 50)
        
    def click_btnCreate_New(self):
        self.pageA.destroy()
        PageB()
        
        

class PageB:

    def __init__(self, master = None):
        self.root = master
        self.pageB = tk.Frame(self.root, width = 1000, height = 700)
        self.pageB.master.title("建立會議")
        self.pageB.grid()
        self.createPage_B()

    def createPage_B(self):
        f1 = tkFont.Font(size = 30, family = "源泉圓體 B")
        f2 = tkFont.Font(size = 20, family = "源泉圓體 M")
        
        global meeting_name
        
        self.lblTitle_B = tk.Label(self.pageB, text = " 創建新會議", height = 1 , width = 15, font = f1, bg = 'DarkSlateGray', fg = 'white', anchor = 'w')
        self.lblname = tk.Label(self.pageB, text = "會議名稱：", height = 1, width = 10, font = f2)
        meeting_name = tk.StringVar()
        self.inputname = tk.Entry(self.pageB, textvariable = meeting_name, width = 20, font = f2)
        # self.txtname = tk.Text(self.pageB, height = 2, width = 40)
        self.lbltime = tk.Label(self.pageB, text = "會議日期：", height = 1, width = 50, font = f2)
        self.btnYes = tk.Button(self.pageB, text = "確認", height = 1, width = 10, font = f2, command = self.click_btnYes)
        
        self.lblTitle_B.place(x = 0, y = 50)
        self.lblname.place(x = 120, y = 150)
        # self.txtname.place(x = 270, y = 150)
        # self.lbltime.place(x = 120, y = 220)
        self.inputname.place(x = 270, y = 150)
        self.btnYes.place(x = 400, y = 450)
        
    def click_btnYes(self):
        self.pageB.destroy()
        # PageA.newMeeting(self, meeting_name.get())
        PageA()


# class PageC:
    # def __init__(self, master = None):
        # self.root = master
        # self.pageC = tk.Frame(self.root, width = 1000, height = 700)
        # self.pageC.master.title("建立會議")
        # self.pageC.grid()
        # self.createPage_C()

    # def createPage_C(self):
        # f1 = tkFont.Font(size = 30, family = "源泉圓體 B")
        # f2 = tkFont.Font(size = 20, family = "源泉圓體 M")
        
        # self.name = meeting_name.get()
        # self.btnCreate_New = tk.Button(self.pageC, text = self.name, height = 1, width = 10, font = f2)
        # self.btnCreate_New.place(x = 30, y = 50)
        # self.btnYes = tk.Button(self.pageC, text = "確認", height = 1, width = 10, font = f2, command = self.click_btnYes)
        # self.btnYes.place(x = 400, y = 450)
    
    # def click_btnYes(self):
        # self.pageC.destroy()
        # # PageA.newMeeting(self, meeting_name.get())
        # PageB()

root.geometry("1000x700") 
root.resizable(0, 0)
Page0(root)

root.mainloop()



