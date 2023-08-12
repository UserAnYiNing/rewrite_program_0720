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
import gc

#變數區-------------------

#廠商－審查順序--------
timeString = '' #時間戳記
fileinput = '' #匯入的申請名單的路徑 / 匯入的抽籤結果的名單的路徑
count_of_data = 0 #資料筆數計算用
record_of_data_amount = 0 #記錄曾匯入過幾筆資料用
applylist = [] #存放申請名單的列表
IDlist_to_draw = [] #記錄treeview的ID的列表
file_before_the_show = '' #未展示的名單的路徑(顯示於申請名單內的，還沒抽中的名單)
file_after_the_show = '' #已展示的名單的路徑(顯示於審查順序名單內的，已經抽中的名單)
#廠商－審查順序--------

#變數區-------------------

#共用事件區-------------------

#檔案轉換用--------

def xlsx_convert_csv(xlsx_path,csv_path): #將xlsx檔轉換成csv檔的副程式
    read_file = pd.read_excel (xlsx_path)
    read_file.to_csv(csv_path, index = None, header=True)

#檔案轉換用--------

#共用事件區-------------------


#廠商－審查順序------------------------------------------------------
def company_sequence():
    sequence_window = tk.Toplevel(index)
    sequence_window.title('高雄市政府青年局抽籤專用軟體')
    sequence_window.geometry('1300x1005+300+0')
    sequence_window.resizable(width=0, height=0)
    sequence_window.after(1,lambda:sequence_window.focus_force())

    # 下面是事件區-------------------

    def on_closing(): #關閉視窗事件
        #將每個使用到的變數初始化

        sequence_window.destroy() #關閉審查順序的視窗

    sequence_window.protocol("WM_DELETE_WINDOW", on_closing) #關閉視窗事件

    def button_event_apply_input():
        global fileinput #連結申請名單檔案路徑的全域變數
        fileinput = filedialog.askopenfilename(parent=sequence_window,title='選擇申請名單的檔案',filetypes=(("Excel Files", "*.xlsx"),)) #開啟選擇檔案的視窗
        if len(fileinput) == 0:
            #如果檢查到路徑字串長度為0的話，就出現錯誤視窗
            msgbox.showerror('錯誤','無選擇檔案',parent=sequence_window)
        else:
            #如果檢查到路徑字串長度不為0的話，便進行檔案讀取
            button_apply_input.configure(state=tk.DISABLED) #將列印按鈕的狀態off，避免誤觸情形發生
            button_to_backup.configure(state=tk.DISABLED) #將備份檔匯入按鈕的狀態off，避免誤觸情形發生

            csv_path = fileinput.replace('.xlsx','.csv') #利用名單檔案路徑字串製作csv檔案路徑
            xlsx_convert_csv(fileinput,csv_path) #使用檔案轉換副程式轉換檔案
            fileinput = csv_path

            file_name = os.path.basename(fileinput).split('.')[0] #取得檔案名稱(不含副檔名)
            filename_text.set(file_name) #將讀取的檔案的名稱顯示在介面上

            global count_of_data #連結資料筆數
            global record_of_data_amount #連結資料總筆數加總數
            if len(table_applylist.get_children()) != 0:
                #再重新匯入申請名單的檔案的話，會把原本的名單記錄清除並換成新匯入的名單資料
                table_applylist.delete(*table_applylist.get_children()) #將申請名單的資料全部清除
                count_of_data = 0 #將名單清除完後，剩餘資料筆數設定為0
                count_of_data_text.set(str(count_of_data)) #將顯示的名單數量更改成更新過後的數字
            if len(table_drawlist.get_children()) != 0:
                # 若是曾啟動過抽獎但資料尚未清除的話，在匯入資料後便會把原資料去除
                table_drawlist.delete(*table_drawlist.get_children()) #將抽籤結果的資料全部清除
                # draw_primarytext.set('')
                # draw_IDtext.set('')
                # draw_taxIDtext.set('')
                # draw_companyname.set('')
            if len(fileinput) != 0:
                with open(fileinput, encoding='UTF-8-sig') as csvFile:
                    csvReader = csv.reader(csvFile) #讀取CSV檔
                    headerRow = next(csvReader)   #讀取計畫名稱
                    projectname_text.set(headerRow[0]) #將顯示計畫名稱用的文字更新
                    headerRow1 = next(csvReader)  # 跳過標題列
                    for row in csvReader:
                        #逐行讀取名單資料
                        count_of_data += 1 #每讀取到一筆即把資料筆數+1
                        applylist.append(row) #將資料加入申請名單列表裡
                        table_applylist.insert("", 'end',text='', values=(row[0],row[1],row[2])) #將資料加進申請名單的treeview
                    count_of_data_text.set(str(count_of_data)) #更新資料筆數的數字

                table_applylist.update() #更新申請名單的treeview
                record_of_data_amount = record_of_data_amount + count_of_data #將現在的資料筆數加進記錄用變數裡，避免多次匯入造成錯誤
                for line in table_applylist.get_children():
                    #記錄列表ID 
                    IDlist_to_draw.append(line) #將讀取到的ID存入列表

    def button_event_backup_input(): #匯入備份
        global fileinput #連結抽籤結果名單的路徑
        global file_before_the_show #連結尚未展示的名單的檔案路徑
        global file_after_the_show #連結已展示的名單的檔案路徑
        fileinput = filedialog.askopenfilename(parent=sequence_window,title='匯入抽籤結果_備份檔',filetypes=(("CSV Files", "*.csv"),))  #取得抽籤結果的檔案路徑
        if len(fileinput) != 0:
            #如果檔案有匯入成功，就進行匯入下一個檔案
            msgbox.showinfo('成功','選擇檔案：' + os.path.basename(fileinput).split('.')[0],parent=sequence_window)
            file_before_the_show = filedialog.askopenfilename(parent=sequence_window,title='匯入未抽到的_審查順序名單_備份檔',filetypes=(("CSV Files", "*.csv"),))  #取得檔案路徑
            if len(file_before_the_show) != 0:
                #如果檔案有匯入成功，就進行匯入下一個檔案
                msgbox.showinfo('成功','選擇檔案：' + os.path.basename(file_before_the_show).split('.')[0],parent=sequence_window)
                file_after_the_show = filedialog.askopenfilename(parent=sequence_window,title='匯入已抽到的_審查順序名單_備份檔',filetypes=(("CSV Files", "*.csv"),))  #取得檔案路徑
                if len(file_after_the_show) != 0:
                    #如果檔案有匯入成功，就進行匯入下一個檔案
                    msgbox.showinfo('成功','選擇檔案：' + os.path.basename(file_after_the_show).split('.')[0],parent=sequence_window)
                    
                    button_apply_input.configure(state=tk.DISABLED) #將列印按鈕的狀態off，避免誤觸情形發生
                    button_to_backup.configure(state=tk.DISABLED) #將備份檔匯入按鈕的狀態off，避免誤觸情形發生
                    
                else:
                    #如果檔案匯入失敗，就顯示訊息視窗，提示使用者需要重新選擇檔案
                    msgbox.showerror('錯誤','無選擇檔案\n請再次點擊備份檔匯入\n重新選擇檔案',parent=sequence_window)
            else:
                #如果檔案匯入失敗，就顯示訊息視窗，提示使用者需要重新選擇檔案
                msgbox.showerror('錯誤','無選擇檔案\n請再次點擊備份檔匯入\n重新選擇檔案',parent=sequence_window)
        else:
            #如果檔案匯入失敗，就顯示訊息視窗，提示使用者需要重新選擇檔案
            msgbox.showerror('錯誤','無選擇檔案\n請再次點擊備份檔匯入\n重新選擇檔案',parent=sequence_window)
    # 上面是事件區-------------------下面是UI設計區


    # 上面是事件區-------------------下面是UI設計區

    #標籤框架用變數
    label_width = 1260

    #標籤框架用變數

    #標籤框架區
    #放按鈕的框格
    buttongroup = tk.LabelFrame(sequence_window,text='', font=('', 14), width=label_width, height=40,borderwidth=0)
    buttongroup.grid(row=0,column=0,padx=20,pady=10)

    #放檔案名稱、計畫名稱和資料筆數顯示用文字的框格
    labelgroup = tk.LabelFrame(sequence_window,text='', font=('', 14), width=label_width, height=90)
    labelgroup.grid(row=1,column=0,padx=20)

    #放抽籤申請列表的框格
    listgroup = tk.LabelFrame(sequence_window,text='', font=('', 14), width=label_width, height=305)
    listgroup.grid(row=2,column=0,padx=20,pady=10)

    #放抽籤結果列表的框格
    resultgroup = tk.LabelFrame(sequence_window,text='', font=('', 14), width=label_width, height=450,bd=1,background='#000')
    resultgroup.grid(row=3,column=0,padx=20)

    bg_img1 = ImageTk.PhotoImage(Image.open(r"data\test_bg.jpg").resize((1255,445)))
    background_label = tk.Label(resultgroup,image=bg_img1)
    background_label.image = bg_img1
    background_label.place(x=0, y=0)
    #標籤框架區

    #功能表(改使用按鈕)
    button_apply_input = tk.Button(buttongroup, text='1.資料匯入', command=button_event_apply_input, font=('DFKai-SB', 18, 'bold'))
    button_apply_input.place(x=1, y=20, anchor="w")
    # button_setting = tk.Button(buttongroup, text='2.設定', command=button_event_setting, font=('DFKai-SB', 18, 'bold'))
    # button_setting.place(x=170, y=20, anchor="w")
    # button_print_announce = tk.Button(buttongroup, text='列印報表', command=button_event_print_all, font=('DFKai-SB', 18, 'bold'),state=tk.DISABLED)
    # button_print_announce.place(x=1120, y=20, anchor="w")
    #功能表(改使用按鈕)

    #匯入用變數
    filename_text = tk.StringVar()  #顯示匯入檔案的名稱用
    filename_text.set('')
    projectname_text = tk.StringVar()  #類別用
    projectname_text.set('')
    count_of_data_text = tk.StringVar()
    count_of_data_text.set(str(count_of_data))
    #匯入用變數

    #抽選標題
    filename_title = tk.Label(labelgroup, text='匯入檔案名稱：', font=('DFKai-SB', 20), bg='#DDD')
    filename_title.place(x=10, y=20, anchor="w") 
    filename_display = tk.Entry(labelgroup, textvariable=filename_text , font=('DFKai-SB', 20), width=54)
    filename_display.place(x=215, y=20, anchor="w") #顯示匯入的檔案的名稱
    projectname_title = tk.Label(labelgroup, text='計畫名稱：', font=('DFKai-SB', 20), bg='#DDD')
    projectname_title.place(x=10, y=60, anchor="w")
    projectname_display = tk.Entry(labelgroup, textvariable=projectname_text ,font=('DFKai-SB', 20), width=77)
    projectname_display.place(x=160, y=60, anchor="w") #顯示計畫的名稱用
    count_of_data_title = tk.Label(labelgroup, text='剩餘筆數：', font=('DFKai-SB', 20), bg='#DDD',fg='#A0A')
    count_of_data_title.place(x=980, y=20, anchor="w")
    count_of_data_display = tk.Entry(labelgroup,textvariable=count_of_data_text , font=('DFKai-SB', 20), width=8, state=tk.DISABLED)
    count_of_data_display.place(x=1130, y=20, anchor="w") #顯示資料剩餘的筆數用的，筆數會隨著抽籤的進行而減少
    #抽選標題

    #申請名單
    apply_label = tk.Label(listgroup, text='申\n請\n名\n單', font=('DFKai-SB', 22), bg='#DDD')
    apply_label.place(x=20, y=150, anchor="w")
    applycolumns = ['案號', '事業名稱','統一編號'] #顯示在表單上的各列欄位名稱
    table_applylist = ttk.Treeview(master=listgroup, height=10, columns=applycolumns, show='headings')
    treeview_style = ttk.Style(master=listgroup)
    treeview_style.configure("Treeview.Heading", font=('DFKai-SB', 22)) #調整所有treeview標題的字體大小
    treeview_style.configure("Treeview", font=('DFKai-SB', 20)) #調整所有treeview內容的字體大小
    treeview_style.configure('Treeview', rowheight=26)

    #下拉bar
    vsb_apply = ttk.Scrollbar(listgroup, orient="vertical", command=table_applylist.yview)
    vsb_apply.place(x=1234, y=6, height=289)
    table_applylist.configure(yscrollcommand=vsb_apply.set)

    table_applylist.heading('案號', text='案號',anchor='nw')  # 定义表头
    table_applylist.heading('事業名稱',text='事業名稱',anchor='nw')
    table_applylist.heading('統一編號', text='統一編號',anchor='nw')  # 定义表头
    table_applylist.column('案號', width=250, minwidth=250,anchor='nw')  # 定义列
    table_applylist.column('事業名稱', width=660, minwidth=660,anchor='nw',)  # 定义列
    table_applylist.column('統一編號', width=270, minwidth=270,anchor='nw')  # 定义列
    table_applylist.place(x=70, y=150, anchor="w")
    #申請名單

    #抽籤顯示結果用text
    draw_primary_text = tk.StringVar()  #顯示順位編號
    draw_primary_text.set('')
    draw_ID_text = tk.StringVar()  #顯示案號
    draw_ID_text.set('')
    draw_taxID_text = tk.StringVar()  #顯示統一編號
    draw_taxID_text.set('')
    draw_companyname_text = tk.StringVar()  #顯示公司名稱
    draw_companyname_text.set('')
    #抽籤顯示結果用text

    #抽籤結果
    draw_primary_title = tk.Label(resultgroup, text='審查順位', font=('DFKai-SB', 32), bg='#DDD')
    draw_primary_title.place(x=20, y=40, anchor="w")
    draw_primary_display = tk.Entry(resultgroup, textvariable=draw_primary_text , font=('Arial', 42, 'bold'), width=8, bg='#E8E39C', fg='#F00')
    draw_primary_display.place(x=210, y=40, anchor="w") #顯示審查順位數字用

    draw_ID_title = tk.Label(resultgroup, text='案號', font=('DFKai-SB', 32), bg='#DDD')
    draw_ID_title.place(x=500, y=40, anchor="w")
    draw_ID_display = tk.Entry(resultgroup, textvariable=draw_ID_text , font=('Arial', 40, 'bold'), width=18, bg='#E8E39C')
    draw_ID_display.place(x=605, y=40, anchor="w") #顯示案號文字用

    draw_taxID_title = tk.Label(resultgroup, text='統一編號', font=('DFKai-SB', 32), bg='#DDD')
    draw_taxID_title.place(x=20, y=190, anchor="w")
    draw_taxID_display = tk.Entry(resultgroup, textvariable=draw_taxID_text , font=('Arial', 40, 'bold'), width=14, bg='#E8E39C')
    draw_taxID_display.place(x=210, y=190, anchor="w") #顯示統一編號文字用


    draw_companyname_title = tk.Label(resultgroup, text='事業名稱', font=('DFKai-SB', 32), bg='#DDD')
    draw_companyname_title.place(x=20, y=115, anchor="w")
    draw_companyname_display = tk.Entry(resultgroup, textvariable=draw_companyname_text , font=('DFKai-SB', 44, 'bold'), width=32, bg='#E8E39C')
    draw_companyname_display.place(x=210, y=115, anchor="w") #顯示事業名稱文字用

    draw_label = tk.Label(resultgroup, text='順\n位\n名\n單', font=('DFKai-SB', 22), bg='#DDD')
    draw_label.place(x=20, y=350, anchor="w")
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
    # button_start_to_draw = tk.Button(sequence_window, text='啟動', command=button_event_starttodraw, font=('DFKai-SB', 22, 'bold'),width=8,background='#F07866')
    # button_start_to_draw.place(x=70, y=960, anchor="w")
    # button_to_clear = tk.Button(sequence_window, text='清除', command=button_event_cleardata, font=('DFKai-SB', 22),width=8,background='#FFF')
    # button_to_clear.place(x=370, y=960, anchor="w")

    button_to_backup = tk.Button(sequence_window, text='備份檔匯入', command=button_event_backup_input, font=('DFKai-SB', 18, 'bold'),width=10)
    button_to_backup.place(x=1000, y=960, anchor="w")
    button_start_to_draw = tk.Button(sequence_window, text='啟動', font=('DFKai-SB', 22, 'bold'),width=8,background='#F07866')
    button_start_to_draw.place(x=70, y=960, anchor="w")
    button_to_clear = tk.Button(sequence_window, text='清除', font=('DFKai-SB', 22),width=8,background='#FFF')
    button_to_clear.place(x=370, y=960, anchor="w")
    #下排按鈕

#廠商－審查順序------------------------------------------------------


#首頁------------------------------------------------------

index = tk.Tk()
index.title('高雄市政府青年局抽籤專用軟體')
index.geometry('490x180+210+210')
index.resizable(width=0, height=0)
index.after(1,lambda:index.focus_force())

def on_closing(): #關閉視窗事件
    gc.collect(generation=2) #在首頁關閉時清除記憶體
    index.destroy() #關閉視窗

index.protocol("WM_DELETE_WINDOW", on_closing) #關閉視窗事件


#功能表
button_company_sequence = tk.Button(index, text='抽廠商之審查順序', command=company_sequence, font=('DFKai-SB', 18, 'bold'),background='#6FD8DE')
button_company_sequence.place(x=10, y=45, anchor="w")
# button_company_primary = tk.Button(index, text='抽廠商正取與備取', command=company_primary, font=('DFKai-SB', 18, 'bold'),background='#D6D5E6')
# button_company_primary.place(x=250, y=45, anchor="w")
# button_people_sequence = tk.Button(index, text='抽民眾之審查順序', command=people_sequence, font=('DFKai-SB', 18, 'bold'),background='#FCDEE4')
# button_people_sequence.place(x=10, y=130, anchor="w")
# button_people_primary = tk.Button(index, text='抽民眾正取與備取', command=people_primary, font=('DFKai-SB', 18, 'bold'),background='#F2E8D5')
# button_people_primary.place(x=250, y=130, anchor="w")
#功能表

index.mainloop()

#首頁------------------------------------------------------