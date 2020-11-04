import pyodbc
import xlrd
import os
from tqdm import tqdm
from tkinter import *
from tkinter import filedialog, ttk

class MyGUI:
    def __init__(self, mainFrame):
        self.window = mainFrame
        self.setupUI()
        self.connect()

    def setupUI(self):
        self.window.title('資料庫匯入精靈')
        windowWidth = 950
        windowHeight = 600
        positionRight = int(self.window.winfo_screenwidth() / 2 - windowWidth / 2)
        positionDown = int(self.window.winfo_screenheight() / 2 - windowHeight / 2)
        self.window.geometry('950x600+{}+{}'.format(positionRight, positionDown))
        self.window.resizable(False, False)

        self.leftFrame = Frame(self.window)
        self.leftFrame.grid(column = 0, row = 0, sticky = 'news')

        self.cbFrame = LabelFrame(self.leftFrame, text = 'Coding Book')
        self.cbFrame.grid(column = 0, row = 0, sticky = 'news')
        self.cbDir = Entry(self.cbFrame)
        self.cbDir.pack(side = 'left', ipadx = 50)
        self.cbBtn = Button(self.cbFrame, text = '選擇檔案', width = 7, height = 2, command = self.chooseCB)
        self.cbBtn.pack(side = 'right')

        self.fileFrame  = LabelFrame(self.leftFrame, text = 'Excel File')
        self.fileFrame.grid(column = 0, row = 1, sticky = 'news')
        self.fileDir = Entry(self.fileFrame)
        self.fileDir.pack(side = 'left', ipadx = 50)
        self.fileBtn = Button(self.fileFrame, text = '選擇檔案', width = 7, height = 2, command = self.chooseFile)
        self.fileBtn.pack(side = 'right')
        
        self.frame = LabelFrame(self.leftFrame, text = '檔案名稱')
        self.frame.grid(column = 0, row = 2, sticky = 'news')
        self.canvas = Canvas(self.frame, height = 400)
        self.canvas.grid(column = 0, row = 0, sticky = 'news', ipadx = 30)
        self.listFrame = Frame(self.canvas)
        self.canvas.create_window(0, 0, window = self.listFrame, anchor = 'nw')
        self.listScroll = Scrollbar(self.frame, orient = VERTICAL, command = self.canvas.yview)
        self.listScroll.grid(column = 1, row = 0, sticky = 'news')
        self.canvas.config(yscrollcommand = self.listScroll.set)
        self.canvas.bind_all("<MouseWheel>", self.scrollCanvas)

        self.btnFrame = Frame(self.leftFrame)
        self.btnFrame.grid(column = 0, row = 3, sticky = W)
        self.createBtn = Button(self.btnFrame, text = '創建', width = 10, command = self.clickCreate)
        self.createBtn.pack(side = 'left', padx = 10, pady = 3)
        self.insertBtn = Button(self.btnFrame, text = '匯入', width = 10, command = self.clickInsert)
        self.insertBtn.pack(side = 'right', padx = 50, pady = 3)
        
        self.rightFrame = Frame(self.window)
        self.rightFrame.grid(column = 1, row = 0, sticky = 'news')

        self.logFrame = LabelFrame(self.rightFrame, text = 'Log')
        self.logFrame.grid(column = 0, row = 0, sticky = 'nw')
        self.logText = Text(self.logFrame)
        self.logText.grid(column = 0, row = 0)
        self.logScroll = Scrollbar(self.logFrame, orient = VERTICAL, command = self.logText.yview)
        self.logScroll.grid(column = 1, row = 0, sticky = 'news')
        self.logText.config(yscrollcommand = self.logScroll.set)
        self.logText.bind_all("<MouseWheel>", self.scrollCanvas)
        self.logText.configure(state = 'disabled')

        self.progressFrame = LabelFrame(self.rightFrame, text = '匯入進度條')
        self.progressFrame.grid(column = 0, row = 1, sticky = 'nws')
        self.progress = ttk.Progressbar(self.progressFrame, orient="horizontal", length = 500, mode="determinate")
        self.progress.grid(column = 0, row = 0, sticky = 'nww')

    def selectFile(self, mode):
        self.selected = {}
        keys = list(self.tablenames.keys())
        values = list(self.tablenames.values())
        for i, var in enumerate(self.vars):
            if var.get() == 1:
                if mode == 1:
                    self.selected[keys[i]] = values[i]
                elif mode == 0:
                    self.selected[values[i]] = self.all_sheet[i]
            
    def clickInsert(self):
        self.selectFile(1)
        for item in self.selected.items():
            self.insert(self.filePath + '/' + item[0], item[1])

    def clickCreate(self):
        self.selectFile(0)
        for item in self.selected.items():
            self.create_table(item[1], item[0])

    def scrollCanvas(self, event):
        self.canvas.yview_scroll(event.delta, "units")

    def setChecklist(self):
        self.vars = []
        self.cbs = []
        row = 1
        self.varAll = IntVar()
        self.cbAll = ttk.Checkbutton(self.listFrame, text = '全選', variable = self.varAll, command = self.selectAll)
        self.cbAll.grid(column = 0, row = 0, sticky = W)

        for i in self.tablenames.keys():
            var = IntVar(0)
            self.vars.append(var)
            c = Checkbutton(self.listFrame, text = i, variable = var)
            self.cbs.append(c)
            c.grid(column = 0, row = row, sticky = W)
            row += 1
        self.listFrame.update_idletasks()
        
        self.canvas.config(scrollregion = self.canvas.bbox('all'))   

    def selectAll(self):
        if self.varAll.get() == 1:
            for cb in self.cbs:
                cb.select()
        else:
            for cb in self.cbs:
                cb.deselect()

    def chooseCB(self):
        self.cbFilename =  filedialog.askopenfilename(initialdir = "~/Downloads",title = "Select file", filetypes = (("excel files","*.xlsx"),("all files","*.*")))
        self.cbDir.insert(0, self.cbFilename)
        # read coding book
        wb = xlrd.open_workbook(self.cbFilename)
        self.all_sheet = wb.sheets()
        sum_table = self.all_sheet[0]
        self.all_sheet = self.all_sheet[1:]
        self.get_filename_tablename(sum_table)
        self.setChecklist()

    def chooseFile(self):
        self.filePath = filedialog.askdirectory(initialdir = "~/Downloads",title = "Select file")
        self.fileDir.insert(0, self.filePath)
        self.test = os.listdir(self.filePath)

    def connect(self):
        self.SERVER = "140.115.105.227"
        self.DB = "IR_EDU"
        UID = "mflab"
        PWD = "dWlab1234%@"
        cn = pyodbc.connect("DRIVER={ODBC Driver 17 for SQL Server};" + "SERVER={};DATABASE={};UID={};PWD={}".format(self.SERVER, self.DB, UID, PWD))
        self.cursor = cn.cursor()

    def create_table(self, sheet, tablename):
        # if table exists
        if self.cursor.tables(table=tablename, tableType='TABLE').fetchone():
            self.logText.configure(state = 'normal')
            self.logText.insert(END, "表{}已存在\n".format(tablename))
            self.logText.configure(state = 'disabled')
            return
        query = ""
        for i in range(2, sheet.nrows):
            # print(tablename, sheet.row_values(i))
            col, dtype = sheet.row_values(i)[1:3]
            if i != 2:
                query += "[" + str(col) + "]" + " " + str(dtype) + ','
            else:
                query += str(col) + " " + str(dtype) + ' NOT NULL IDENTITY(1,1) PRIMARY KEY,'
                
        query = query.strip(',')

        query = "create table {}({});".format(tablename, query)
        try:
            self.logText.configure(state = 'normal')
            self.cursor.execute(query)
            self.cursor.commit()
            self.logText.insert(END, "創建{}完成".format(tablename) + '\n')
            self.logText.configure(state = 'disabled')
        except Exception as e:
            self.logText.configure(state = 'normal')
            self.logText.insert(END, '創建{}發生錯誤 : '.format(tablename) + str(type(e)) + " " +  str(e) + '\n')
            self.logText.insert(END, '錯誤的SQL query : ' + query + '\n')
            self.logText.configure(state = 'disabled')

    def insert(self, file, tablename):
        progressValue = 0
        error = False
        dataType = []
        for row in self.cursor.columns(table=tablename):
            dataType.append(row[5])
        dataType = dataType[1:]
        # print(dataType)
        wb = xlrd.open_workbook(file)
        sheet = wb.sheets()[0]

        for i in range(1, sheet.nrows):
            progressValue = float(i / sheet.nrows) * 100
            self.progress["value"] = progressValue
            self.progress.update()
            tmp = sheet.row_values(i)
            tmp = tmp[:sheet.ncols]
            query = ""
            for idx, j in enumerate(tmp):
                if isinstance(j, float):
                    j = str(int(j))
                if "\'" in j:
                    j = j.replace("\'", "\'\'")

                if dataType[idx] == 'int' or dataType[idx] == 'tinyint':
                    if j == '':
                        j = "\'\'"
                    query += j + ','
                else:
                    query += "\'" + j + "\'" + ','
                
            query = query.strip(',')
            try:
                self.cursor.execute("insert into {}.dbo.{} values(".format(self.DB, tablename) + query + ");")
                
            except Exception as e:
                self.logText.configure(state = 'normal')
                self.logText.insert(END, "匯入{}發生錯誤 : ".format(tablename) + str(type(e)) + " " + str(e) + '\n')
                self.logText.insert(END, "錯誤的SQL query : " + query + '\n')
                self.logText.configure(state = 'disabled')
                error = True
                break
        if not error:
            self.cursor.commit()
            self.progress["value"] = 500
            self.progress.update()
            self.logText.configure(state = 'normal')
            self.logText.insert(END, "匯入{}完成".format(tablename) + '\n')
            self.logText.configure(state = 'disabled')
    def get_filename_tablename(self, summary_table):
        self.tablenames = {}
        for row in range(1, summary_table.nrows):
            filename, tablename = summary_table.row_values(row)[4:6]
            filename += '.xls'
            self.tablenames[filename] = tablename
        
    



if __name__ == "__main__":
    window = Tk()
    gui = MyGUI(window)

    # create tables
    # for i, tablename in enumerate(tablenames.values()):
    #     create_table(all_sheet[i], tablename)

    # for name in tablenames.items():
    #     prefix = './20200917/校庫合併_公開資訊9910_10903/'
    #     insert(prefix + name[0], name[1])

    # prefix = './20200917/校庫合併_公開資訊9910_10903/'
    
    # insert(prefix + '研4. 學校學術研究計畫成效表.xls', "R_Research_Project")

    window.mainloop()