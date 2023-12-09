# -*- coding: utf-8 -*-
from  tkinter import *
from  tkinter import ttk, messagebox, filedialog
import tkinter as tk
import pandas as pd
import sys


from tkcalendar import DateEntry
import numpy as np
from datetime import datetime
import os

# 列印PDF套件
# import pdfkit
# from ironpdf import *
import asyncio
from pyppeteer import launch

# from PyQt5.QtGui import *
# from PyQt5.QtWidgets import *
# from PyQt5.QtCore import *
# from test import DateTimeEditDemo

global data, search_data
global path1, file_input, file_input_path
# 資料來源目錄
if getattr(sys, 'frozen', False): 
    PATH = os.path.dirname(sys.executable)
elif __file__: 
    PATH = os.path.dirname(__file__) 

# 資料來源目錄
path1 = f"{PATH}/SourceData"
file_input = "TreeView_SourceData.xlsx"
file_input_path = os.path.join(path1, file_input)
# print(file_input_path)

def Dataloader(path1, file_input_path):
    if os.path.exists(path1) == False:
        try:
            os.mkdir(path1)
            print("目錄建立成功")
            if os.path.exists(file_input_path) == False:
                newdata = pd.DataFrame([],columns=['id', 'company','description', 'quantity', 'un', 'amount','date','mark'])
                newdata.to_excel(file_input_path, index=False)
                print("檔案新增成功")
        except:
            print("目錄建立失敗,請手動建立SourceData資料夾")
    else:
        if os.path.exists(file_input_path) == False:
            newdata = pd.DataFrame([],columns=['id', 'company','description', 'quantity', 'un', 'amount','date','mark'])
            print(newdata)
            newdata.to_excel(file_input_path, index=False)
            print("檔案新增成功")    
                   
    df = pd.read_excel(file_input_path, parse_dates=['date'])
    df_removenan = df.replace(np.nan, "")
    # df_removenan["date"] = df_removenan['year'].map(str)+"-"+df_removenan['month'].map(str)+"-"+df_removenan['day'].map(str)
    df_removenan["quantity"] = df_removenan["quantity"].map(lambda x: f"{x:,}")
    df_removenan["un"] = df_removenan["un"].map(lambda x: f"{x:,}")
    df_removenan["amount"] = df_removenan["amount"].map(lambda x: f"{x:,} 元")
    # df_removenan.astype({'date':"datetime64[ns]"})
    # df_removenan['date'] = pd.to_datetime(df_removenan['date'])
    # df_removenan['date'] =df_removenan['date'].dt.strftime('%Y-%m-%d')
    return df_removenan
    
data = Dataloader(path1, file_input_path)
search_data = data

##########################################
# 初始進入基準介面
class basedesk():
    def __init__(self, master):
        self.root = master
        self.root.config()
        width = root.winfo_screenwidth()
        height = root.winfo_screenheight()
        root.geometry(f"{width}x{height}")  # width x height
        root.title("進出貨料")  # Adding a title    
        initface(self.root)

##########################################
# 進貨資料(查詢資料頁面)
class initface():
    def back(self):
        self.face1.destroy()
        initface(self.master)

    def __init__(self, master):
        for widget in root.winfo_children():
            widget.destroy()
        self.master = master
        data = Dataloader(path1, file_input_path)
        # 基准界面initface
        self.initface = tk.Frame(self.master, )
        self.initface.pack()
        menubar = tk.Menu(root)
       
        editmenu = tk.Menu(menubar, tearoff=False)
        editmenu.add_command(label="查詢進貨資料", command= self._page_search)
        editmenu.add_command(label="新增進貨資料", command= self._page_add)
        editmenu.add_command(label="離開", command=root.quit)
        menubar.add_cascade(label="選擇模式（查詢/新增與刪除）", menu=editmenu)
        root.config(menu = menubar)
        
        frame_all = tk.Frame(root, borderwidth=2, relief="groove")
        frame_all.pack(ipadx=5, ipady=5,padx=10, pady=10, fill='x')
        frame_fix = tk.Frame(frame_all)
        frame_fix.pack(ipady=2,padx=2, pady=5, fill='x')

        data_columns = list(data.columns)  # columns=['orderid', 'company','description', 'quantity', 'un', 'amount','date','mark']
        l1=list(data)
        r_set=data.to_numpy().tolist() # Create list of list using rows 
        name = ['編號', '寶號','品名', '數量', '單價', '金額','日期','備註']
        # print(len(name)) #10
        
        # orderid
        lab_orderid  = tk.Label(frame_fix, text=name[0]+str(": "), font=("Helvetica",18))
        lab_orderid.pack(side='left', anchor='nw',pady=1,padx=2)
        txtbox_orderid = tk.Text(frame_fix, font=("Helvetica",18), background='white',border=2,height=1, width=15)
        txtbox_orderid.pack(side='left',anchor='nw',pady=1,padx=1,expand=1)
        # company
        lab_company  = tk.Label(frame_fix, text=name[1]+str(": "), font=("Helvetica",18))
        lab_company.pack(side='left', anchor='nw',pady=1,padx=2)
        txtbox_company = tk.Text(frame_fix, font=("Helvetica",18), background='white',height=1, width=30)
        txtbox_company.pack(side='left',anchor='nw',pady=1,padx=1,expand=1)

        frame_choose = tk.Frame(frame_all)
        frame_choose.pack(ipady=2,padx=2, pady=5, fill='x')
        # description
        lab_description  = tk.Label(frame_choose, text=name[2]+str(": "), font=("Helvetica",18))
        lab_description.pack(side='left',anchor='nw',pady=1,padx=1,expand=0)
        txtbox_description = tk.Text(frame_choose, font=("Helvetica",18), background='white',height=1, width=30)
        txtbox_description.pack(side='left',anchor='nw',pady=1,padx=1,expand=1)
        # quantity
        #Create a Calendar using DateEntry
        lab_description  = tk.Label(frame_choose, text=name[6]+str(": "), font=("Helvetica",18))
        lab_description.pack(side='left',anchor='nw',pady=1,padx=1,expand=0)
        cal_start = DateEntry(frame_choose,bootstyle='primary',font=('Helvetica',18))
        cal_start.pack(side='left',anchor='nw',pady=1,padx=1,expand=0)
        lab2_description  = tk.Label(frame_choose, text=' - ', font=("Helvetica",18))
        lab2_description.pack(side='left',anchor='nw',pady=1,padx=1,expand=0)
        cal_end = DateEntry(frame_choose,bootstyle='primary',font=('Helvetica',18))
        cal_end.pack(side='left',anchor='nw',pady=1,padx=1,expand=1)




        # app = QApplication(sys.argv)
        # demo = DateTimeEditDemo()
        # demo.show()
        # sys.exit(app.exec_())

        # btn_PDF=tk.Button(frame_choose,text='另存資料(備份)',font=('Helvetica',18), fg='black',width=10,borderwidth=2,command=lambda:research())
        # btn_PDF.pack(ipadx=30,padx=20,pady=5, ipady=10, side='right',anchor='e',expand=0)


        frame_search_mode = tk.Frame(frame_all)
        frame_search_mode.pack(ipady=2,padx=2, pady=5, fill='x')

        mylabel = tk.Label(frame_search_mode, font=('Arial',30), fg='#f00')  # 放入標籤
        mylabel.pack(side='left')
            
        val_id = tk.BooleanVar()
        check_btn_ID = tk.Checkbutton(frame_search_mode, text='依編號查詢', variable=val_id, onvalue=bool(True), 
                                      offvalue=bool(False), state=tk.DISABLED)
        check_btn_ID.pack(side='left',anchor='w',pady=1,padx=1,expand=0)
        check_btn_ID.deselect()
        
        val_company = tk.StringVar()
        check_btn_company = tk.Checkbutton(frame_search_mode, text='依寶號查詢', variable=val_company, onvalue='company',
                                           offvalue='', state=tk.DISABLED)
        check_btn_company.pack(side='left',anchor='w',pady=1,padx=1,expand=0)
        check_btn_company.deselect()
        
        val_description = tk.StringVar()
        check_btn_description = tk.Checkbutton(frame_search_mode, text='依品名查詢', variable=val_description, onvalue='description', 
                                               offvalue='', state=tk.DISABLED)
        check_btn_description.pack(side='left',anchor='w',pady=1,padx=1,expand=0)
        check_btn_description.deselect()
        
        val_time = tk.StringVar()
        check_btn_time = tk.Checkbutton(frame_search_mode, text='依時間查詢', variable=val_time, onvalue='time',
                                        offvalue='')
        check_btn_time.select()
        check_btn_time.pack(side='left',anchor='w',pady=1,padx=1,expand=0)
        
        
        btn_search = tk.Button(frame_search_mode,
                        text='查詢',
                        font=('Helvetica',24,'bold'),command=lambda:cal_search_data()
                    )
        btn_search.pack(ipadx=60,padx=20,pady=20, ipady=10,side='left',expand=0)
        
        def cal_search_data():
            global search_data
            search_data_start_date = cal_start.get()
            search_data_end_date =  cal_end.get()
            # print(search_data_start_date)
            # print(search_data_end_date)
            
            # Convert the date to datetime64
            # data['date'] = pd.to_datetime(data['date'], format='%Y-%m-%d')
            data['date'] = pd.to_datetime(data['date'],format='mixed', dayfirst=True)
            

            # Filter data between two dates
            search_data = data.loc[(data['date'] >= search_data_start_date) & (data['date'] <= search_data_end_date)]
            # print(search_data)
            
            for i in tree.get_children():
                tree.delete(i)
            search_data_list = search_data.to_numpy().tolist()
            for itm in search_data_list:
                tree.insert("",END,values=itm)
        
        # btn_PDF=tk.Button(frame_search_mode,text='另存資料',font=('Helvetica',18), fg='black',command=lambda:research())
        # btn_PDF.pack(ipadx=30,padx=20,pady=5, ipady=10,anchor='e',expand=0)
        # frame_search_mode.pack(padx=2, pady=2,fill='x')

        btn_PDF=tk.Button(frame_search_mode,text='重新查詢',font=('Helvetica',18), fg='black',width=10,borderwidth=2,command=lambda:research())
        btn_PDF.pack(ipadx=30,padx=20,pady=5, ipady=10,anchor='e',expand=0)
        frame_search_mode.pack(padx=2, pady=2,fill='x')
        
        def research():
            for i in tree.get_children():
                tree.delete(i)
            search_data_list = data.to_numpy().tolist()
            for itm in search_data_list:
                tree.insert("",END,values=itm)

        btn_PDF=tk.Button(frame_search_mode,text='輸出PDF',font=('Helvetica',18), fg='#3366ff',width=10,borderwidth=2, command=lambda:gen_pdf())
        btn_PDF.pack(ipadx=30,padx=20,pady=5, ipady=10, side='right',anchor='e',expand=0)
        # 列印查詢結果至PDF檔
        def gen_pdf():
            options = {
                'page-size': 'A4',
                'margin-top': '0.75in',
                'margin-right': '0.75in',
                'margin-bottom': '0.75in',
                'margin-left': '0.75in',
                'encoding': "UTF-8",
                'custom-header': [
                    ('Accept-Encoding', 'gzip')
                ],
                'no-outline':None}
            html_string = '''
                <!DOCTYPE html>
                <html>
                    <head><title>進貨資料查詢</title>
                    <h1>進貨資料查詢</h1>
                    <style>
                        table, td, th {{border: 1px solid black;}}
                        table {{border-collapse: collapse; width: 100%;}}
                        td, th {{padding: 0px 10px 0px 5px;}}
                        th {{text-align: center;}}
                        td {{text-align: center;}}
                        h1 {{text-align: center;}}
                    </style>
                    </head>
                    <body style="background-color:red">
                        {table}
                    </body>
                </html>
            '''
            fileName =filedialog.asksaveasfilename(defaultextension=".pdf",filetypes=[("PDF",".pdf"),("HTML",".html"),("JPG",".jpg")])
            file, ext = os.path.splitext(fileName)
            search_data_print = search_data
            search_data_print.rename(columns={"id":"訂單編號", "company":"寶號", "description":"品名", "quantity":"數量", "un":"單價", "amount":"總價", "date":"交易日期", "mark": "備註"}, inplace=True)
            search_data_print = search_data_print.reset_index(drop=True)
            async def generate_pdf_from_html(html_content, pdf_path):
                browser = await launch()
                page = await browser.newPage()
                
                await page.setContent(html_content)
                
                await page.pdf({'path': pdf_path, 'format': 'A4'})
                
                await browser.close()
            print(fileName)
            try:
                asyncio.get_event_loop().run_until_complete(generate_pdf_from_html(html_string.format(table=search_data_print.to_html()), fileName))
                messagebox.showinfo('showinfo', message='檔案儲存成功')
            except:
                messagebox.showerror('showerror', message='檔案儲存失敗') 


            # asyncio.get_event_loop().run_until_complete(generate_pdf(html_string.format(table=search_data_print.to_html()), fileName))

            # if(ext[1:]=='pdf'):
                
                ###### pdf-ironpdf ######
                # renderer = ChromePdfRenderer()
                # pdf = renderer.RenderHtmlAsPdf(html_string.format(table=search_data_print.to_html()))
                # pdf.SaveAs(fileName)
                ###### pdf-pdfkit ######
                # print_test = pdfkit.from_string(html_string.format(table=search_data_print.to_html()), fileName, options=options)

                # messagebox.showinfo('showinfo', message='檔案儲存成功') 
            
            # search_path = f"{PATH}//SearchResult"
            # HTML_path = os.path.join(search_path, f'{datetime.now().strftime("%Y%m%d_%H%M%S")}.html')
            # PDF_path = os.path.join(search_path, f'{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf')

            # search_data.to_html(HTML_path)
            
            
            # if os.path.exists(search_path) == False:
            #     os.mkdir(search_path)
            
            # try:
            #     search_data.to_html(HTML_path)
            #     # pdfkit.from_file(HTML_path, PDF_path, configuration=config, options=options)
            #     pdfkit.from_file(HTML_path, PDF_path, options=options)
            #     messagebox.showinfo('showinfo', message='條件設定搜尋資料\nPDF檔輸出成功\nPDF檔案路徑: '+str(PDF_path))     
            # except:
            #     messagebox.showerror('showerror', 'showerror')
            # # pdfkit.from_url('http://google.com', 'out.pdf', configuration=config)
        



        # 列表資料
        # ldata = [(20230101, "嘉義股份有限公司", "線材", format(1000, ","), format(20, ","), format(20000, ",") ), 
        #         (20230101, "嘉義股份有限公司", "影印紙", format(15, ","), format(15, ","), format(1500, ",")),
        #         (20230101, "嘉義股份有限公司", "亮面影印紙", format(25, ","), format(100, ","), format(2500, ","))]
        scrollbar = tk.Scrollbar(root)
        scrollbar.pack(side="right", fill="y")
        tree = ttk.Treeview(root, columns=(data_columns), show="headings", displaycolumns="#all", yscrollcommand=scrollbar.set)
        for i in range(len(data_columns)):
            tree.heading(data_columns[i], text=name[i])
        
        # # datalist = list(data)
        # print(data)
        tree.pack(expand=1,side='left', fill='x')
        for itm in r_set:
            tree.insert("",END,values=itm)
        tree.pack(expand=1, fill=BOTH)
        tree.column('0', minwidth=100, width=150, stretch=False)
        tree.column('1', minwidth=150, width=150, stretch=False)
        tree.column('2', minwidth=150, width=200, stretch=False)
        tree.column('3', minwidth=50, width=80, stretch=False, anchor='e')
        tree.column('4', minwidth=50, width=80, stretch=False, anchor='e')
        tree.column('5', minwidth=200, width=200, stretch=False, anchor='e')
        tree.column('6', minwidth=150, width=150, stretch=False, anchor='c')
        scrollbar.config(command=tree.yview)


          
        
    def _page_add(self, ):
        self.initface.destroy()
        face1(self.master) # 新增資料頁面
        
    def _page_search(self, ):
        self.initface.destroy()
        initface(self.master)


##########################################
# 進貨資料(新增資料頁面)
n = 1 
class face1():
    def back(self):
            self.face1.destroy()
            initface(self.master)

    def __init__(self, master):
        for widget in root.winfo_children():
            widget.destroy()
        self.master = master
        # self.master.config(bg='gray')
        # 基准界面initface
        self.initface = tk.Frame(self.master, )
        self.initface.pack()
        
        menubar = tk.Menu(root)
        editmenu = tk.Menu(menubar, tearoff=False)
        editmenu.add_command(label="查詢進貨資料", command= self._page_search)
        editmenu.add_command(label="新增進貨資料", command= self._page_add)
        editmenu.add_command(label="離開", command=root.quit)
        menubar.add_cascade(label="選擇模式（查詢/新增與刪除）",font=("Helvetica",24), menu=editmenu)
        root.config(menu = menubar)
        
        frame_add_fix = tk.Frame(root)
        frame_add_fix.pack()
        label_add_fix = tk.Label(frame_add_fix, text='新增進貨資料', font=("Helvetica",28))
        label_add_fix.pack(anchor="center",pady=10,padx=10)
        
        frame_add_table = tk.Frame(root)
        frame_add_table.pack(fill='both', anchor='center')
        columns = ('0', '1', '2', '3', '4', '5', '6', '7')
        tree = ttk.Treeview(frame_add_table, show='headings', columns=columns)
        tree.column('0', width=20, anchor='center')
        tree.column('1', anchor='w')
        tree.column('2', anchor='center')
        tree.column('3', width=60, anchor='e')
        tree.column('4', width=80, anchor='e')
        tree.column('5', anchor='w')
        tree.column('6', anchor='w')
        tree.column('7', anchor='w')
        
        # tree.tag_configure('+', background='red')
        # tree.tag_configure('-', background='green')
        tree.heading('0', text='#')
        tree.heading('1', text='日期')
        tree.heading('2', text='寶號')
        tree.heading('3', text='品項')
        tree.heading('4', text='數量')
        tree.heading('5', text='單價')
        tree.heading('6', text='總價')
        tree.heading('7', text='備註')
        tree.pack(padx=20,side='left',expand=0, anchor='w', pady=10)
        
        ######## 選取新增的資料，同步顯示至輸入的方框 ########
        def treeSelect(event):
            item = tree.selection()
            itemvalues = tree.item(item, 'values')

            '''
            測試用 看有幾筆資料
            print(len(tree.get_children()))
            測試用 看item值
            print(len(itemvalues),itemvalues)
            n=0
            for i in itemvalues:
                print('item{}={}'.format(n,i))
                n+=1
            '''
            if (len(itemvalues)>0):
                # 清除輸入框
                clearEntry()
                # 更新輸入框的值
                cal_input.insert(0, itemvalues[1])
                txt_input_company.insert(0, itemvalues[2])
                txt_input_description.insert(0, itemvalues[3])
                txt_input_quantity.insert(0, itemvalues[4])
                txt_input_un.insert(0, itemvalues[5])
                txt_input_amount.insert(0, itemvalues[6])
                txt_input_mark.insert(0, itemvalues[7])
        tree.bind('<<TreeviewSelect>>', treeSelect)
                
        # def modify_column():
        #     pass
        
        ######## 刪除新增的資料 ########
        def del_column():
            try:
                selected_item = tree.selection()[0]  # get selected item
                tree.delete(selected_item)
            except:
                messagebox.showwarning("錯誤訊息", "請點選要刪除的資料!!!")
        
        ######## 儲存新增的資料 ########
        def savefile():
            global n
            n = 1
            # with open("mystock_price.csv", "w", newline='') as myfile:
            #     csvwriter = csv.writer(myfile, delimiter=',')
            Sdatetime = datetime.now().strftime("%Y%m%d%H%M%S")
            tmp_data = []
            for row_id in tree.get_children():
                row = tree.item(row_id)['values']
                # print('save row:', row)
                del row[0:1]
                row.insert(0,Sdatetime)
                tmp_data.append(row)
            
            test_dataframe = pd.DataFrame(tmp_data, columns=["id", "date", "company", "description", "quantity", "un", "amount", "mark"])
            inser_data_to_excel = test_dataframe[["id", "company", "description", "quantity", "un", "amount", "date", "mark"]]

            file_out = file_input_path
            # 如果檔案不存在，建立檔案
            # 如果檔案存在，append            
            if len(inser_data_to_excel)==0:
                messagebox.showerror(title='資料新增失敗', message='您未輸入任何資料，請正確輸入！')
            else:
                if os.path.exists(file_out) == False:
                    inser_data_to_excel.to_excel(file_out,index=False)
                elif  os.path.exists(file_out) == True:
                    df_old = pd.read_excel(file_out)
                    df_combine = df_old._append(inser_data_to_excel)
                    df_combine.to_excel(file_out,index=False)
                messagebox.showinfo(title='資料新增成功', message='資料新增成功')
                for i in tree.get_children():
                    tree.delete(i)

                    # messagebox.showerror(title='資料新增失敗', message='資料新增失敗')


#                try:
#                    if os.path.exists(file_out) == False:
#                        inser_data_to_excel.to_excel(file_out,index=False)
#                    elif  os.path.exists(file_out) == True:
#                        df_old = pd.read_excel(file_out)
#                        df_combine = df_old.append(inser_data_to_excel)
#                        df_combine.to_excel(file_out,index=False)
#                    messagebox.showinfo(title='資料新增成功', message='資料新增成功')
#                    for i in tree.get_children():
#                        tree.delete(i)
#                except:
#                    messagebox.showerror(title='資料新增失敗', message='資料新增失敗')
            
        btn_save = tk.Button(frame_add_table, text='儲存資料', font=("Helvetica",18),fg='red', command=savefile)
        btn_save.pack(padx=5,ipadx=3, ipady=5, fill='x')
        
        # btn_rewrite = tk.Button(frame_add_table, text='修改資料', font=("Helvetica",18), command=modify_column)
        # btn_rewrite.pack(padx=5,ipadx=3, ipady=3, fill='x')
        
        btn_delet = tk.Button(frame_add_table, text='刪除資料', font=("Helvetica",18), command=del_column)
        btn_delet.pack(padx=5,ipadx=3, ipady=3, fill='x')

    
        
        frame_add_change = tk.Frame(root,height=20)
        frame_add_change.pack(padx=10,fill='both')
        
        lbl_date = tk.Label(frame_add_change,font=('Arial', 14), text="日期:", fg="green")
        lbl_date.pack(side='left',expand=0, anchor='w')
        # cal_input = DateEntry(frame_add_change,
        #         fieldbackground='light green',
        #         background='dark green',
        #         foreground='dark blue',
        #         arrowcolor='white')
        cal_input = DateEntry(frame_add_change,bootstyle='primary',font=('Arial', 14))
        cal_input.pack(side='left',padx=10, pady=10, expand=1, anchor='w')
        # cal = DateEntry(frame_choose, width= 16, background= "magenta3", foreground= "gray",bd=2)
        # cal.pack(side='left',anchor='nw',pady=1,padx=1,expand=0)
        # txt_input_date = tk.Entry(frame_add_change, width=10)
        # txt_input_date.pack(side='left', expand=1, anchor='w')
        
        lbl_company = tk.Label(frame_add_change,font=('Arial', 14), text="寶號:", fg="green")
        lbl_company.pack(side='left',expand=0, anchor='w')
        txt_input_company = tk.Entry(frame_add_change,font=('Arial', 14))
        txt_input_company.pack(side='left', expand=1, anchor='w')
        
        lbl_description = tk.Label(frame_add_change,font=('Arial', 14), text="品名:", fg="black")
        lbl_description.pack(side='left',expand=0, anchor='w')
        txt_input_description = tk.Entry(frame_add_change,font=('Arial', 14))
        txt_input_description.pack(side='left', expand=1, anchor='w')
        
        lbl_quantity = tk.Label(frame_add_change,font=('Arial', 14), text="數量:", fg="black")
        lbl_quantity.pack(side='left',expand=0, anchor='w')
        txt_input_quantity = tk.Spinbox(frame_add_change,font=('Arial', 14), from_=1, to=999999, width=8)
        txt_input_quantity.pack(side='left', expand=1, anchor='w')
        
        lbl_un = tk.Label(frame_add_change,font=('Arial', 14), text="單價:", fg="black")
        lbl_un.pack(side='left',expand=0, anchor='w')
        txt_input_un = tk.Entry(frame_add_change,font=('Arial', 14), width=8)
        txt_input_un.pack(side='left', expand=1, anchor='w')
        
        lbl_amount = tk.Label(frame_add_change,font=('Arial', 14), text="總價:", fg="black")
        lbl_amount.pack(side='left',expand=0, anchor='w')
        txt_input_amount = tk.Entry(frame_add_change,font=('Arial', 14), width=8)
        txt_input_amount.pack(side='left', expand=1, anchor='w')

        lbl_mark = tk.Label(frame_add_change,font=('Arial', 14), text="備註:", fg="black")
        lbl_mark.pack(side='left',expand=0, anchor='w')
        txt_input_mark = tk.Entry(frame_add_change,font=('Arial', 14))
        txt_input_mark.pack(side='left', expand=1, anchor='w')
    
    
    
        # 新增資料至 tree_view 中
        frame_btn = tk.Frame(root)
        frame_btn.pack(pady=3)

        def InputData():
            global n
            addDate = cal_input.get()
            addCompany = txt_input_company.get()
            addDescription = txt_input_description.get()
            addquantity = txt_input_quantity.get()
            addUn = txt_input_un.get()
            addAmount = txt_input_amount.get()
            addMark = txt_input_mark.get()
            if addDate == '' or addCompany == '' or addDescription == '' or addquantity == '' or addUn =='':
                messagebox.showwarning("錯誤訊息", "請輸入資料!!!")
            else:
                if (not addquantity.isdigit()) or (not addUn.isdigit()) or (not addAmount.isdigit):
                    messagebox.showwarning("錯誤訊息", "請檢查輸入資料格式是否正確!!! \n (數量、單價與總價須為數值格式)")
                else:
                    i = [str(n), str(addDate), str(addCompany), str(addDescription), int(addquantity), int(addUn), int(addAmount), str(addMark)]
                    tree.insert('', 'end', values=i)
                    n += 1
                    # cal_input.delete(0, 'end')
                    # txt_input_company.delete(0, 'end')
                    txt_input_description.delete(0, 'end')
                    txt_input_quantity.delete(0, 'end')
                    txt_input_un.delete(0, 'end')
                    txt_input_amount.delete(0, 'end')
                    txt_input_mark.delete(0, 'end')
        
        def clearEntry():
                cal_input.delete(0, 'end')
                txt_input_company.delete(0, 'end')
                txt_input_description.delete(0, 'end')
                txt_input_quantity.delete(0, 'end')
                txt_input_un.delete(0, 'end')
                txt_input_amount.delete(0, 'end')
                txt_input_mark.delete(0, 'end')
                
        btn_confirm = tk.Button(frame_add_change, text='新增資料', font=("Helvetica",18),fg='green', command=InputData)
        btn_confirm.pack(ipadx=3, ipady=3,side='left', expand=1, anchor='w')
      
        btn_clear = tk.Button(frame_add_change, text='清除資料', font=("Helvetica",18), command=clearEntry)
        btn_clear.pack(padx=5, ipadx=3, ipady=3,side='left', expand=1, anchor='w')
        
    def _page_add(self, ):
        self.initface.destroy()
        face1(self.master) # 新增資料頁面
        
    def _page_search(self, ):
        self.initface.destroy()
        initface(self.master)
    

         

if __name__ == '__main__':
    root = tk.Tk()
    basedesk(root)    
    root.mainloop()
