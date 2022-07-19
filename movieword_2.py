from cProfile import label
from enum import auto
import openpyxl
import tkinter as tk

from pyparsing import col

win = tk.Tk()
win.geometry("800x300")
win.title("熊的單字庫")

Ewordafter = tk.StringVar()
Cwordafter = tk.StringVar()
FromName = tk.StringVar()
Ewordbefore = tk.StringVar()
workbook = openpyxl.load_workbook('moviewo.xlsx')
count=len(workbook.sheetnames)

def searchEword():
    EwordafterString = ""
    CwordafterString = ""
    FromNameString = ""
    for k in range(0,count):
            sheet = workbook.worksheets[k]
            for i in range(1,sheet.max_row+1):
                for j in range(1,sheet.max_column+1):
                    if Ewordbefore.get() == str(sheet.cell(row=i,column=j).value):
                        EwordafterString = EwordafterString + "\n" + str(sheet.cell(row=i,column=j).value)
                        CwordafterString = CwordafterString + "\n" + str(sheet.cell(row=i,column=j+1).value)
                        FromNameString = FromNameString + "\n" + str(workbook.sheetnames[k])
                        break
    Ewordafter.set(EwordafterString)
    Cwordafter.set(CwordafterString)
    FromName.set(FromNameString)
    EwordafterString = ""
    CwordafterString = ""
    FromNameString = ""

frame1 = tk.Frame(win)
frame1.pack()
label1 = tk.Label(frame1, text="想查詢的單字：", font=("微軟正黑體", 18))
entry = tk.Entry(frame1, textvariable=Ewordbefore, font=("微軟正黑體", 18))
labelNone1 = tk.Label(frame1, text=" ", font=("微軟正黑體", 18))
button1 = tk.Button(frame1, text="查詢", font=("微軟正黑體", 18), fg="blue", command=searchEword)
label1.grid(row=0, column=0)
entry.grid(row=0, column=1)
labelNone1.grid(row=0, column=2)
button1.grid(row=0, column=3)

labelNone2 = tk.Label(win, text="")
labelNone2.pack()

frame2 = tk.Frame(win)
frame2.pack()
labeltitle1 = tk.Label(frame2, text="英文單字".ljust(15), font=("微軟正黑體", 18), bg="black", fg="white")
labeltitle2 = tk.Label(frame2, text="中文".ljust(60), font=("微軟正黑體", 18), bg="black", fg="white")
labeltitle3 = tk.Label(frame2, text="出處".ljust(30), font=("微軟正黑體", 18), bg="black", fg="white")
label2 = tk.Label(frame2, textvariable=Ewordafter, font=("微軟正黑體", 14))
label3 = tk.Label(frame2, textvariable=Cwordafter, font=("微軟正黑體",14))
label4 = tk.Label(frame2, textvariable=FromName, font=("微軟正黑體", 14))
labeltitle1.grid(row=0, column=0)
labeltitle2.grid(row=0, column=1)
labeltitle3.grid(row=0, column=2)
label2.grid(row=1, column=0, sticky="w")
label3.grid(row=1, column=1, sticky="w")
label4.grid(row=1, column=2, sticky="w")

win.mainloop()