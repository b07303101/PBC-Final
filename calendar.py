import calendar  # 引用日曆套件
import tkinter as tk  
import tkinter.font as tkFont  # 調整字體
from tkinter import ttk  # 美化元件外觀
import openpyxl

datetime = calendar.datetime.datetime #日期和時間結合
timedelta = calendar.datetime.timedelta  # 時間差
class Calendar:
    def __init__(s, point = None, position = None):
        # point    窗口位置
        # position 窗口在點的位置 'ur'-右上, 'ul'-左上, 'll'-左下, 'lr'-右下
        s.master = tk.Toplevel()  # 多一個視窗
        fwday = calendar.SUNDAY 

        year = datetime.now().year  # 打開頁面時為當下年月份
        month = datetime.now().month
        locale = None
        sel_bg = '#ecffc4'  #選取框框色
        sel_fg = '#05640e'  # 選取字底色

        s._date = datetime(year, month, 1)  # 該月份第一天
        s._selection = None # 設置未選中的日期

        s.G_Frame = ttk.Frame(s.master)

        s._cal = s.__get_calendar(locale, fwday)   # 實例化適當的日曆類

        s.__setup_styles()       # 創建自定義樣式
        s.__place_widgets()      # pack/grid 小部件
        s.__config_calendar()    # 調整日曆列和安裝標記
        # 配置畫布和正確的绑定，以選擇日期。
        s.__setup_selection(sel_bg, sel_fg)
        # 存儲項ID，用於稍後插入。
        s._items = [s._calendar.insert('', 'end', values='') for _ in range(6)]

        # 在當前空日曆中插入日期
        s._update()

        s.G_Frame.pack(expand = 1, fill = 'both')
        s.master.overrideredirect(1)  # 調整字
        s.master.update_idletasks()  # 刷新頁面
        width, height = s.master.winfo_reqwidth(), s.master.winfo_reqheight()  # 寬高居中
        if   position == 'ur':
            x, y = point[0], point[1] - height  
        elif position == 'lr':
            x, y = point[0], point[1]
        elif position == 'ul':
            x, y = point[0] - width, point[1] - height
        elif position == 'll':
            x, y = point[0] - width, point[1]
        else:
            x, y = (s.master.winfo_screenwidth() - width)/2, (s.master.winfo_screenheight() - height)/2
        s.master.geometry('%dx%d+%d+%d' % (width, height, x, y)) #窗口位置居中
        s.master.after(300, s._main_judge)  #每300秒更新一次
        s.master.deiconify()  # 還原視窗 
        s.master.focus_set()  # 焦點設置在所需小部件上
        s.master.wait_window() 
        
        
    def __get_calendar(s, locale, fwday):  # 日曆文字化
        if locale is None:
            return calendar.TextCalendar(fwday)
        else:
            return calendar.LocaleTextCalendar(fwday, locale)
            
    def __setup_styles(s): # 自定義TTK風格
        style = ttk.Style(s.master)
        arrow_layout = lambda dir: ([('Button.focus', {'children': [('Button.%sarrow' % dir, None)]})])  # 返回參數性質
        style.layout('L.TButton', arrow_layout('left'))  #製作點選上個月的箭頭
        style.layout('R.TButton', arrow_layout('right'))  # 製作點選下個月的箭頭
        
    
    def __place_widgets(s):  # 標題框架及其小部件
        Input_judgment_num = s.master.register(s.Input_judgment)  # 需要将函数包装一下，必要的
        hframe = ttk.Frame(s.G_Frame)
        gframe = ttk.Frame(s.G_Frame)
        bframe = ttk.Frame(s.G_Frame)
        hframe.pack(in_=s.G_Frame, side='top', pady=5, anchor='center')  # 日曆的上視窗
        gframe.pack(in_=s.G_Frame, fill=tk.X, pady=5)
        bframe.pack(in_=s.G_Frame, side='bottom', pady=5)  

        lbtn = ttk.Button(hframe, style='L.TButton', command=s._prev_month)  # 左箭頭
        lbtn.grid(in_=hframe, column=0, row=0, padx=12)
        rbtn = ttk.Button(hframe, style='R.TButton', command=s._next_month)  # 右箭頭
        rbtn.grid(in_=hframe, column=5, row=0, padx=12)
        
        s.CB_year = ttk.Combobox(hframe, width = 5, values = [str(year) for year in range(datetime.now().year, datetime.now().year+11,1)], validate = 'key', validatecommand = (Input_judgment_num, '%P'))  # 製作下拉選單
        s.CB_year.current(0)  # 下拉式索引起初在該年分
        s.CB_year.grid(in_=hframe, column=1, row=0)
        s.CB_year.bind("<<ComboboxSelected>>", s._update)
        tk.Label(hframe, text = '年', justify = 'left').grid(in_=hframe, column=2, row=0, padx=(0,5))  # 下拉式選單後面的單位(年)

        s.CB_month = ttk.Combobox(hframe, width = 3, values = ['%02d' % month for month in range(1,13)], state = 'readonly')  # 下拉式選單-月
        s.CB_month.current(datetime.now().month - 1)
        s.CB_month.grid(in_=hframe, column=3, row=0)
        s.CB_month.bind("<<ComboboxSelected>>", s._update)
        tk.Label(hframe, text = '月', justify = 'left').grid(in_=hframe, column=4, row=0)  # 下拉式選單後面的單位(年)
        # 日曆部件
        s._calendar = ttk.Treeview(gframe, show='', selectmode='none', height=7)
        s._calendar.pack(expand=1, fill='both', side='bottom', padx=5)
        ttk.Button(bframe, text = "加入", width = 6, command = lambda: s._exit(True)).grid(row = 0, column = 0, sticky = 'ns', padx = 20)
        ttk.Button(bframe, text = "取消", width = 6, command = s._exit).grid(row = 0, column = 1, sticky = 'ne', padx = 20)
        tk.Frame(s.G_Frame, bg = '#565656').place(x = 0, y = 0, relx = 0, rely = 0, relwidth = 1, relheigh = 2/200)
        tk.Frame(s.G_Frame, bg = '#565656').place(x = 0, y = 0, relx = 0, rely = 198/200, relwidth = 1, relheigh = 2/200)
        tk.Frame(s.G_Frame, bg = '#565656').place(x = 0, y = 0, relx = 0, rely = 0, relwidth = 2/200, relheigh = 1)
        tk.Frame(s.G_Frame, bg = '#565656').place(x = 0, y = 0, relx = 198/200, rely = 0, relwidth = 2/200, relheigh = 1)
        
    def __config_calendar(s):  # 設計日曆架構
        cols = ['日','一','二','三','四','五','六']  # 日曆上的星期幾
        s._calendar['columns'] = cols  # 設定日曆欄
        s._calendar.tag_configure('header', background='grey90')
        s._calendar.insert('', 'end', values=cols, tag='header')  # 調整其列寬
        font = tkFont.Font()
        maxwidth = max(font.measure(col) for col in cols)
        for col in cols:
            s._calendar.column(col, width=maxwidth, minwidth=maxwidth,anchor='center')

    def __setup_selection(s, sel_bg, sel_fg):
        def __canvas_forget(evt):
            canvas.place_forget()
            s._selection = None

        s._font = tkFont.Font()
        s._canvas = canvas = tk.Canvas(s._calendar, background=sel_bg, borderwidth=0, highlightthickness=0)
        canvas.text = canvas.create_text(0, 0, fill=sel_fg, anchor='w')

        
        s._calendar.bind('<Button-1>', s._pressed)  #點出要的日期
  
        
    def _build_calendar(s):
        year, month = s._date.year, s._date.month  # update header text (Month, YEAR)
        header = s._cal.formatmonthname(year, month, 0)  # 更新日曆顯示的日期
        cal = s._cal.monthdayscalendar(year, month)
        for indx, item in enumerate(s._items):
            week = cal[indx] if indx < len(cal) else []
            fmt_week = [('%02d' % day) if day else '' for day in week]
            s._calendar.item(item, values=fmt_week)
            
    def _show_select(s, text, bbox):  # 秀出挑選的日子
        x, y, width, height = bbox

        textw = s._font.measure(text)
        canvas = s._canvas
        canvas.configure(width = width, height = height)
        canvas.coords(canvas.text, (width - textw)/2, height / 2 - 1)
        canvas.itemconfigure(canvas.text, text=text)
        canvas.place(in_=s._calendar, x=x, y=y)
        
    def _pressed(s, evt = None, item = None, column = None, widget = None):  #在日曆的某個地方點擊。
        
        if not item:
            x, y, widget = evt.x, evt.y, evt.widget
            item = widget.identify_row(y)
            column = widget.identify_column(x)
        if not column or not item in s._items:  # 在工作日行中單擊或僅在列外單擊。
            return
        item_values = widget.item(item)['values']  # 點選的日期該周
        if not len(item_values): # 該行是空的。
            return
        text = item_values[int(column[1]) - 1] # text 為選擇日期
        if not text: # 日期為空
            return
        bbox = widget.bbox(item, column)
        if not bbox: # 日曆尚未出現
            s.master.after(20, lambda : s._pressed(item = item, column = column, widget = widget))
            return
        text = '%02d' % text
        s._selection = (text, item, column)
        s._show_select(text, bbox)
    def _prev_month(s):
        s._canvas.place_forget()
        s._selection = None

        s._date = s._date - timedelta(days=1)
        s._date = datetime(s._date.year, s._date.month, 1)
        s.CB_year.set(s._date.year)
        s.CB_month.set(s._date.month)
        s._update()

    def _next_month(s):
        s._canvas.place_forget()
        s._selection = None

        year, month = s._date.year, s._date.month
        s._date = s._date + timedelta(
            days=calendar.monthrange(year, month)[1] + 1)
        s._date = datetime(s._date.year, s._date.month, 1)
        s.CB_year.set(s._date.year)
        s.CB_month.set(s._date.month)
        s._update()

    def _update(s, event = None, key = None):
        if key and event.keysym != 'Return': return
        year = int(s.CB_year.get())
        month = int(s.CB_month.get())
        if year == 0 or year > 9999: return
        s._canvas.place_forget()
        s._date = datetime(year, month, 1)
        s._build_calendar() # 重建日曆

        if year == datetime.now().year and month == datetime.now().month:
            day = datetime.now().day
            for _item, day_list in enumerate(s._cal.monthdayscalendar(year, month)):
                if day in day_list:
                    item = 'I00' + str(_item + 2)
                    column = '#' + str(day_list.index(day)+1)
                    s.master.after(100, lambda :s._pressed(item = item, column = column, widget = s._calendar))

    def _exit(s, confirm = False):
        if not confirm:
            s._selection = None
        s.master.destroy()

    def _main_judge(s):
        try:
            #s.master 為 TK 窗口
            #if not s.master.focus_displayof(): s._exit()
            #else: s.master.after(10, s._main_judge)

            #s.master 為 toplevel 窗口
            if s.master.focus_displayof() == None or 'toplevel' not in str(s.master.focus_displayof()): s._exit()
            else: s.master.after(10, s._main_judge)
        except:
            s.master.after(10, s._main_judge)

        #s.master.tk_focusFollowsMouse() # 焦點跟随游標
    def selection(s):
        if not s._selection: 
            return None
        year = s._date.year
        month = s._date.month
        return str(datetime(year, month, int(s._selection[0])))[:10]

    def Input_judgment(s, content):
        if content.isdigit() or content == "":
            return True
        else:
            return False

if __name__ == '__main__':  
    date_list =[]
    root = tk.Tk()
    width = root.winfo_reqwidth() + 50
    height = 100 #窗口大小
    x, y = (root.winfo_screenwidth()  - width )/2, (root.winfo_screenheight() - height)/2
    root.geometry("1000x700") #窗口位置居中
    date_str = tk.StringVar()
    date = ttk.Entry(root, textvariable = date_str)  # 放選擇的日期
    date.place(x = 250, y = 100, relx = 1/20, rely = 1/6, relwidth = 4/8, relheigh = 2/30)
    date_str_gain = lambda: [date_str.set(date) for date in [Calendar((x, y), 'ur').selection()] if date]
    def press_add():
        date_list.append(str(date.get()))
        for i in range(len(date_list)):
            date_label = ttk.Label(root,text = "你已選擇"+date_list[i]+",")
        date_label.place(x = 300, y = 150, relx = 1/20, rely = 1/6, relwidth = 4/8, relheigh = 2/30)
        date.delete(0,"end")
    def write():
        press_add()
        wb = openpyxl.load_workbook('C:/Users/sunny/Desktop/日期.xlsx')
        sheet = wb['工作表1']
        for i in range(len(date_list)):
            sheet.cell(row=1,column=i+2).value = date_list[i]
        wb.save('C:/Users/sunny/Desktop/日期.xlsx')
    tk.Button(root, text = '選擇日期:', command = date_str_gain).place(x = 200, y = 100, relx = 1/20, rely = 1/6, relwidth = 4/80, relheigh = 2/30)
    tk.Button(root, text = "新增日期", command = press_add).place(x = 350, y = 300, relx = 1/20, rely = 1/6, relwidth = 4/80, relheigh = 2/30)
    tk.Button(root, text = "確定",command = write).place(x = 450, y = 300, relx = 1/20, rely = 1/6, relwidth = 4/80, relheigh = 2/30)
    root.mainloop()
   


        

