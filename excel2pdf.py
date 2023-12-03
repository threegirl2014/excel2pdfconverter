#!Python 3
#网上找的代码，适当改动
#源代码依然保留如下，按excel确认xls与xlsm文件所属位置，pdf确认文件输出路径，单击第三个按钮完成转换。
#在excel下（包括子文件夹）都会被转化
import sys
import os
import win32com.client
import tkinter.filedialog
from os.path import splitext
from tkinter import *
import tkinter.messagebox

pdf_path=None

def callback1():
    global excel_path
    excel_path = tkinter.filedialog.askdirectory()
    print('已确认excel文件所属位置' + excel_path)

def callback2():
    global pdf_path
    pdf_path = tkinter.filedialog.askdirectory()
    print('已确认pdf文件输出位置' + pdf_path)

def callback3():
    for root, dirs, files in os.walk(excel_path):
        for file in files:
            if file.endswith('.xls') or file.endswith('.xlsx'):
                print('处理excel文件：' + file)
                if file.startswith('~$'):
                    print('该文件无法处理，直接跳过！')
                    continue
                file1 = splitext(file)
                if pdf_path is None:
                    out_file_path = root + '/' + file1[0]
                else:
                    out_file_path = pdf_path + '/' + file1[0]
                filename = os.path.join(root, file)
                in_file = os.path.abspath(filename)
                out_file = os.path.abspath(out_file_path)
                o = win32com.client.Dispatch("Excel.Application")
                o.Visible=False
                try:
                    wb = o.Workbooks.Open(in_file)
                    # wb.ActiveSheet.ExportAsFixedFormat(0, out_file+'.pdf')
                    wb.ExportAsFixedFormat(0, out_file+'.pdf')
                    print('转换pdf完成：' + out_file + '.pdf')
                    wb.Close(SaveChanges=0)
                except Exception as e:
                    print('发生异常' + str(e))
                    

if __name__=="__main__":
    root = Tk()
    root.title('GD_excel转pdf_v1.0')
    Button(root, text="EXCEL路径", fg="blue",bd=2,width=28,command=callback1).pack()
    Button(root, text="PDF输出路径", fg="blue",bd=2,width=28,command=callback2).pack()
    Button(root, text="转换到PDF输出路径", fg="blue",bd=2,width=28,command=callback3).pack()
    Button(root, text="EXCEL原路径一键输出", fg="red", bd=2,width=28,command=callback3).pack()
    root.mainloop()
