from tkinter import *
# import  tkinter
# 导入ttk
from tkinter import ttk
# 导入simpledialog
from tkinter import simpledialog
from tkinter import filedialog

import Docxfile

class View:
    numat = 0
    filepath =''
    status = 0
    subkey = ''
    doc =''
    savename=''
    def __init__(self, master):
        self.master = master
        self.text = Text(self.master)
        # self.createDoc()
        self.initWidgets()

    def initWidgets(self):
        keyarray = ['a','b', 'c', 'd', 'e', 'f', 'g', 'i', 'j','k']

        self.text.config(wrap=WORD)
        self.text.insert(1.0,'Running... ')
        self.text.insert(END, '\n')
        self.text.pack()

        self.open_integer()

        self.open_filename()

        self.doc = Docxfile.Docxfile(self.filepath)
        self.doc.genDocument()

        for ik in range(self.numat):
            key = '<@'+keyarray[ik]+'>'
            # print(key)
            self.open_string(key)

            self.callReplace(key,self.subkey)

        self.savename = simpledialog.askstring("保存文件名为：", "不包含后缀(默认为docx文件)",
                                             initialvalue='')
        self.doc.saveFile(self.savename)
        self.text.insert(END, '文件名字保存为：'+self.savename)


        self.text.insert(END, '任务结束。')
        self.text.insert(END, '\n')
        self.text.pack()

    def createDoc(self):
        self.doc = Docxfile.Docxfile(self.filepath)
        self.doc.genDocument()


    def callReplace(self, key, subkey):
        if self.doc.replaceContent(key, subkey) == 0:
            self.text.insert(END, '已完成。')
            self.text.insert(END, '\n')
            self.text.pack()


    def open_string(self,key):
        # 调用askstring函数生成一个让用户输入字符串的对话框
        self.subkey = simpledialog.askstring("替换", "便签"+key,
                                     initialvalue='')
        self.text.insert(END, '替换标签'+key+' 为 ： '+self.subkey)
        self.text.insert(END, '\n')
        self.text.insert(END, '替换ing....')
        self.text.insert(END, '\n')
        self.text.pack()

    def open_integer(self):

        # 调用askinteger函数生成一个让用户输入整数的对话框
        self.numat = simpledialog.askinteger("修改数", "模版文件中有多少<@>:",
                                      initialvalue=1, minvalue=1, maxvalue=10)
        if self.numat <= 0:
            self.text.insert(END, '模版文件中<@>数量: ')
            self.text.insert(END, 'Error: Need to be larger than 0.')
            self.status = -1;
        else:
            self.text.insert(END, '模版文件中<@>数量: ')
            self.text.insert(END, self.numat)

        self.text.insert(END, '\n')
        self.text.pack()



    def open_filename(self):
        # 调用askopenfilename方法获取单个文件的文件名
        self.filepath = filedialog.askopenfilename(title='打开单个文件',
                filetypes=[("Microsoft Word document", "*.docx")],  # 只处理的文件类型
                initialdir='/Users/kent/PycharmProjects/autoDocx/')  # 初始目录

        if self.filepath.strip()=="":
            self.text.insert(END, '模版文件路径： ')
            self.text.insert(END, 'Error: Need to correct path.')
            self.status = -1;
        else:
            self.text.insert(END, '模版文件路径： ')
            self.text.insert(END, self.filepath)

        self.text.insert(END, '\n')
        self.text.pack()

root = Tk()
root.title("autoDocx")
View(root)
root.mainloop()
