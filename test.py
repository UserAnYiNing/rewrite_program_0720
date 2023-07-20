import tkinter as tk  #UI用
from tkinter import ttk, messagebox as msgbox, filedialog #UI用
import csv #匯入用/備份用
import random #亂數用
import time #間隔時間、時間戳記用
import os
import docx #word檔案編輯用
from docx.shared import Pt, Cm #word內容格式用
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn #word內容格式用
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT # 导入单元格垂直对齐
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT # 导入段落对齐
from PIL import ImageTk, Image
from docx2pdf import convert
import subprocess
import sys
import shutil  #複製備份檔用
import pandas as pd
import openpyxl

#共同使用
record = 1 #記錄曾匯入過幾筆資料
datacounter = 0 #連接到匯入和輸入 計算匯入資料筆數用
fileinput = '' #一般資料(順序/正備取共用)匯入路徑
file_applyinput = '' #備份資料(申請)(順序/正備取共用)匯入路徑
applylist = []  #儲存尚未抽取到的名單
IDlisttodraw = []  #儲存名單ID
#共同使用

#審查順序用
drawresultlist = []  #抽籤後依抽籤順位排列審查順序的資料用
savepath_pdf = ''  #pdf檔案儲存路徑
savepath_word = '' #word檔案儲存路徑
file_drawinput = '' #備份資料(結果)匯入路徑
#審查順序用

#正備取用
primary_counter = 0 #記錄抽取正取資料筆數用
reverse_counter = 0 #記錄抽取備取資料筆數用
file_primary_backup = '' #備份資料(正取結果)匯入路徑
file_reverse_backup = '' #備份資料(備取結果)匯入路徑
primary_result = [] #抽籤後依抽籤順位排列正取的資料用
reverse_result = [] #抽籤後依抽籤順位排列備取的資料用
savepath_primary_pdf = ''  #pdf正取檔案儲存路徑
savepath_reverse_pdf = ''  #pdf備取檔案儲存路徑
savepath_primary_word = '' #word正取檔案儲存路徑
savepath_reverse_word = '' #word備取檔案儲存路徑
primary_amount_all = 0 #匯入正取備份資料時用來記錄檔案內資料數量
reverse_amount_all = 0 #匯入備取備份資料時用來記錄檔案內資料數量
#正備取用

#檔案轉換用
def xlsx_convert_csv(xlsx_path,csv_path):
    read_file = pd.read_excel (xlsx_path)
    read_file.to_csv (csv_path, index = None, header=True)
#檔案轉換用

#匯入用變數
label1_text = tk.StringVar()  #名稱用
label1_text.set('')
label3_text = tk.StringVar()  #類別用
label3_text.set('')
#匯入用變數

#抽籤顯示結果用text
draw_primarytext = tk.StringVar()  #顯示順位編號
draw_primarytext.set('')
draw_IDtext = tk.StringVar()  #顯示案號
draw_IDtext.set('')
draw_taxIDtext = tk.StringVar()  #顯示統一編號
draw_taxIDtext.set('')
draw_companyname = tk.StringVar()  #顯示公司名稱
draw_companyname.set('')
timeString = '' #時間戳記
#抽籤顯示結果用text

#廠商－審查順序------------------------------------------------------
def company_sequence():
    root = tk.Toplevel(index)
    root.title('高雄市政府青年局抽籤專用軟體')
    root.geometry('1300x1005+300+0')
    root.resizable(width=0, height=0)
    root.after(1,lambda:root.focus_force())
    
    # -------------------

    # -------------------

    def button_event_input():
        pass
    
    def button_event_setting():
        pass

    def button_event_print_all():
        pass

    def button_event_starttodraw():
        pass

    def button_event_cleardata():
        pass

    def button_event_backup_input():
        pass
    
    # 上面是事件區-------------------下面是UI設計區




    # 上面是事件區-------------------下面是UI設計區

    #標籤框架用變數
    label_width = 1260

    #標籤框架用變數

    #標籤框架區
    buttongroup = tk.LabelFrame(root,text='', font=('', 14), width=label_width, height=40,borderwidth=0)
    buttongroup.grid(row=0,column=0,padx=20,pady=10)

    labelgroup = tk.LabelFrame(root,text='', font=('', 14), width=label_width, height=90)
    labelgroup.grid(row=1,column=0,padx=20)

    listgroup = tk.LabelFrame(root,text='', font=('', 14), width=label_width, height=305)
    listgroup.grid(row=2,column=0,padx=20,pady=10)

    resultgroup = tk.LabelFrame(root,text='', font=('', 14), width=label_width, height=450,bd=1,background='#000')
    resultgroup.grid(row=3,column=0,padx=20)

    back_img1 = ImageTk.PhotoImage(Image.open(r"data\test_bg.jpg").resize((1255,445)))
    background_label = tk.Label(resultgroup,image=back_img1)
    background_label.image = back_img1
    background_label.place(x=0, y=0)

    #標籤框架區

    #功能表(改使用按鈕)
    button_print_input = tk.Button(buttongroup, text='1.資料匯入', command=button_event_input, font=('DFKai-SB', 18, 'bold'))
    button_print_input.place(x=1, y=20, anchor="w")
    button_setting = tk.Button(buttongroup, text='2.設定', command=button_event_setting, font=('DFKai-SB', 18, 'bold'))
    button_setting.place(x=170, y=20, anchor="w")
    button_print_announce = tk.Button(buttongroup, text='列印報表', command=button_event_print_all, font=('DFKai-SB', 18, 'bold'),state=tk.DISABLED)
    button_print_announce.place(x=1120, y=20, anchor="w")
    #功能表(改使用按鈕)

    #抽選標題
    label1 = tk.Label(labelgroup, text='匯入檔案名稱：', font=('DFKai-SB', 20), bg='#DDD')
    label1.place(x=10, y=20, anchor="w")
    label2 = tk.Entry(labelgroup, textvariable=label1_text , font=('DFKai-SB', 20), width=54)
    label2.place(x=215, y=20, anchor="w")
    label3 = tk.Label(labelgroup, text='計畫名稱：', font=('DFKai-SB', 20), bg='#DDD')
    label3.place(x=10, y=60, anchor="w")
    label4 = tk.Entry(labelgroup, textvariable=label3_text ,font=('DFKai-SB', 20), width=77)
    label4.place(x=160, y=60, anchor="w")
    label5 = tk.Label(labelgroup, text='剩餘筆數：', font=('DFKai-SB', 20), bg='#DDD',fg='#A0A')
    label5.place(x=980, y=20, anchor="w")
    label6_text = tk.StringVar()
    label6_text.set(str(datacounter))
    label6 = tk.Entry(labelgroup,textvariable=label6_text , font=('DFKai-SB', 20), width=8, state=tk.DISABLED)
    label6.place(x=1130, y=20, anchor="w")
    #抽選標題

    #申租名單
    mylabel = tk.Label(listgroup, text='申\n請\n名\n單', font=('DFKai-SB', 22), bg='#DDD')
    mylabel.place(x=20, y=150, anchor="w")
    listcolumns = ['案號', '事業名稱','統一編號']
    table_applylist = ttk.Treeview(master=listgroup, height=10, columns=listcolumns, show='headings')
    treeview_style = ttk.Style(master=listgroup)
    treeview_style.configure("Treeview.Heading", font=('DFKai-SB', 22)) #調整所有treeview標題的字體大小
    treeview_style.configure("Treeview", font=('DFKai-SB', 20)) #調整所有treeview內容的字體大小
    treeview_style.configure('Treeview', rowheight=26)

    #下拉bar
    vsb_apply = ttk.Scrollbar(listgroup, orient="vertical", command=table_applylist.yview)
    vsb_apply.place(x=1234, y=35, height=260)
    table_applylist.configure(yscrollcommand=vsb_apply.set)

    table_applylist.heading('案號', text='案號',anchor='nw')  # 定义表头
    table_applylist.heading('事業名稱',text='事業名稱',anchor='nw')
    table_applylist.heading('統一編號', text='統一編號',anchor='nw')  # 定义表头
    table_applylist.column('案號', width=250, minwidth=250,anchor='nw')  # 定义列
    table_applylist.column('事業名稱', width=660, minwidth=660,anchor='nw',)  # 定义列
    table_applylist.column('統一編號', width=270, minwidth=270,anchor='nw')  # 定义列
    table_applylist.place(x=70, y=150, anchor="w")
    #申租名單

    #抽籤結果
    mylabel1 = tk.Label(resultgroup, text='審查順位', font=('DFKai-SB', 32), bg='#DDD')
    mylabel1.place(x=20, y=40, anchor="w")
    mylabel1_1 = tk.Entry(resultgroup, textvariable=draw_primarytext , font=('Arial', 42, 'bold'), width=8, bg='#E8E39C', fg='#F00')
    mylabel1_1.place(x=210, y=40, anchor="w")

    mylabel2 = tk.Label(resultgroup, text='案號', font=('DFKai-SB', 32), bg='#DDD')
    mylabel2.place(x=500, y=40, anchor="w")
    mylabel2_1 = tk.Entry(resultgroup, textvariable=draw_IDtext , font=('Arial', 40, 'bold'), width=18, bg='#E8E39C')
    mylabel2_1.place(x=605, y=40, anchor="w")

    mylabel3 = tk.Label(resultgroup, text='統一編號', font=('DFKai-SB', 32), bg='#DDD')
    mylabel3.place(x=20, y=190, anchor="w")
    mylabel3_1 = tk.Entry(resultgroup, textvariable=draw_taxIDtext , font=('Arial', 40, 'bold'), width=14, bg='#E8E39C')
    mylabel3_1.place(x=210, y=190, anchor="w")


    mylabel3_2 = tk.Label(resultgroup, text='事業名稱', font=('DFKai-SB', 32), bg='#DDD')
    mylabel3_2.place(x=20, y=115, anchor="w")
    mylabel3_3 = tk.Entry(resultgroup, textvariable=draw_companyname , font=('DFKai-SB', 44, 'bold'), width=32, bg='#E8E39C')
    mylabel3_3.place(x=210, y=115, anchor="w")

    mylabel4 = tk.Label(resultgroup, text='順\n位\n名\n單', font=('DFKai-SB', 22), bg='#DDD')
    mylabel4.place(x=20, y=350, anchor="w")
    drawcolumns = ['審查順位','案號', '事業名稱', '統一編號']
    table_drawlist = ttk.Treeview(master=resultgroup, height=7, columns=drawcolumns, show='headings',selectmode='none')

    vsb_draw = ttk.Scrollbar(resultgroup, orient="vertical", command=table_drawlist.yview)
    vsb_draw.place(x=1234, y=259, height=182)
    table_drawlist.configure(yscrollcommand=vsb_draw.set)

    table_drawlist.heading('審查順位', text='審查順位',anchor='n')  # 定义表头
    table_drawlist.heading('案號',text='案號',anchor='nw')
    table_drawlist.heading('事業名稱', text='事業名稱',anchor='nw')  # 定义表头
    table_drawlist.heading('統一編號', text='統一編號',anchor='nw')  # 定义表头

    table_drawlist.column('審查順位', width=150, minwidth=150,anchor='n')  # 定义列
    table_drawlist.column('案號', width=250, minwidth=250,anchor='nw')  # 定义列
    table_drawlist.column('事業名稱', width=530, minwidth=530,anchor='nw')  # 定义列
    table_drawlist.column('統一編號', width=250, minwidth=250,anchor='nw')  # 定义列
    table_drawlist.place(x=70, y=335, anchor="w")
    #抽籤結果

    #下排按鈕
    mybutton_start = tk.Button(root, text='啟動', command=button_event_starttodraw, font=('DFKai-SB', 22, 'bold'),width=8,background='#F07866')
    mybutton_start.place(x=70, y=960, anchor="w")
    mybutton_clear = tk.Button(root, text='清除', command=button_event_cleardata, font=('DFKai-SB', 22),width=8,background='#FFF')
    mybutton_clear.place(x=370, y=960, anchor="w")

    mybutton_backup = tk.Button(root, text='備份檔匯入', command=button_event_backup_input, font=('DFKai-SB', 18, 'bold'),width=10)
    mybutton_backup.place(x=1000, y=960, anchor="w")
    #下排按鈕
    
#廠商－審查順序------------------------------------------------------

#廠商－正(備)取------------------------------------------------------
def company_primary():
    pass
#廠商－正(備)取------------------------------------------------------

#民眾－審查順序------------------------------------------------------
def people_sequence():
    pass
#民眾－審查順序------------------------------------------------------

#民眾－正(備)取------------------------------------------------------
def people_primary():
    pass
#民眾－正(備)取------------------------------------------------------

#首頁

index = tk.Tk()
index.title('高雄市政府青年局抽籤專用軟體')
index.geometry('490x180+210+210')
index.resizable(width=0, height=0)
index.after(1,lambda:index.focus_force())


#功能表
button_print_input = tk.Button(index, text='抽廠商之審查順序', command=company_sequence, font=('DFKai-SB', 18, 'bold'),background='#6FD8DE')
button_print_input.place(x=10, y=45, anchor="w")
button_print_setting = tk.Button(index, text='抽廠商正取與備取', command=company_primary, font=('DFKai-SB', 18, 'bold'),background='#D6D5E6')
button_print_setting.place(x=250, y=45, anchor="w")
mybutton_clear = tk.Button(index, text='抽民眾之審查順序', command=people_sequence, font=('DFKai-SB', 18, 'bold'),background='#FCDEE4')
mybutton_clear.place(x=10, y=130, anchor="w")
button_print_announce = tk.Button(index, text='抽民眾正取與備取', command=people_primary, font=('DFKai-SB', 18, 'bold'),background='#F2E8D5')
button_print_announce.place(x=250, y=130, anchor="w")
#功能表

index.mainloop()

#首頁
