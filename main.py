#!/usr/bin/evn python
#-*- coding:utf-8 -*-
'''
Created on 2018年6月20日

@author: jinfeng
'''
import tkinter as tk  
from tkinter import ttk  
from tkinter import scrolledtext  
from tkinter import Menu  
from tkinter import Spinbox  
from tkinter import messagebox as mBox
from tkinter.filedialog  import asksaveasfilename 
import os
import time
from connect_to_mssql import *
from write_xls import *

host = '127.0.0.1'
user = 'sa'
password = '12345678'
database = 'test'
conn = connect_to_mssql(host,user,password,database)

def input_infor():
    global conn,display_vartext
    list_value = []
    for e in list_entry:
        if e.get() != '':
            list_value.append(e.get())
        else:
            list_value.append('Null')
        e.set('')
    print(list_value)
    insert_value(conn,list_value)
    display_vartext.set('输入信息已保存！')
    mBox.showinfo("messagebox","this is button 4 dialog")  

def select_infor():
    global conn
    condition1 = "%s like '%%%s%%'" % (label2_1_stringvar.get(),label2_1_stringvar2.get())
    condition2 = "%s like '%%%s%%'" % (label2_2_stringvar.get(),label2_2_stringvar2.get())
    condition3 = "%s like '%%%s%%'" % (label2_3_stringvar.get(),label2_3_stringvar2.get())
    condition = condition1 + ' and ' + condition2 + ' and ' + condition3
    commit = 'select * from informationtable where %s' % condition
    result = send_select(conn,commit)
    items = tree.get_children()
    for id in items:
        tree.delete(id) #删除原有数据
    for i in range(len(result)):
        tree.insert("",i,values=result[i]) #插入数据
        
def save_to_xls():
    values = []
    filename=asksaveasfilename(parent = windows,defaultextension='.xls')
    items = tree.get_children()
    for id in items:
        values.append(tree.item(id,'values'))
    write_xls(list_label,values,filename)


windows = tk.Tk()
windows.geometry('645x560')
windows.resizable(False, False)
windows.title("人员信息管理系统  Version 1.0.6 ")

# 插入Notebook
tabControl = ttk.Notebook(windows)
  
tab1 = ttk.Frame(tabControl)
tabControl.add(tab1, text='录入')
  
tab2 = ttk.Frame(tabControl)
tabControl.add(tab2, text='查询') 
  
tab3 = ttk.Frame(tabControl)  
tabControl.add(tab3, text='修改')   

tabControl.pack(expand=1, fill="both")  

# Notebook 第一页面控件
lf1 = ttk.LabelFrame(tab1, text='信息录入')
lf1.grid(column=0, row=0, padx=8, pady=4)  

n = 0
list_label = get_rowname(conn)
list_entry = []
if len(list_label)%2 == 0:
    tag_number = len(list_label)//2
else:
    tag_number = len(list_label)//2+1
while n<len(list_label):
    vartext = tk.StringVar()
    list_entry.append(vartext)
    if n < tag_number:
        label = ttk.Label(lf1,text=list_label[n])
        label.grid(row=n,column=0, padx=6,pady=10, sticky=tk.W)
        entry = tk.Entry(lf1,width=30,textvariable=list_entry[n])
        entry.grid(row=n,column=1,padx=6,pady=10, sticky=tk.W)
    else:
        label = ttk.Label(lf1,text=list_label[n])
        label.grid(row=n-tag_number,column=2,padx=6,pady=10, sticky=tk.W)
        entry = tk.Entry(lf1,width=30,textvariable=list_entry[n])
        entry.grid(row=n-tag_number,column=3,padx=6,pady=10, sticky=tk.W)
    #list_entry.append(entry)
    n += 1
    
button_input = tk.Button(lf1,width=12,text="录入",command=input_infor)
button_input.grid(row=n+1,column=1, padx=8, pady=4)
display_vartext = tk.StringVar()
display_vartext.set('请在文本框内输入信息，点击“录入”按钮保存至数据库。')
dispaly_label = tk.Label(lf1,textvariable=display_vartext,fg='red')
dispaly_label.grid(row=n+1,column=2,columnspan=2, padx=8, pady=4)

# Notebook 第二页面控件
lf2 = ttk.LabelFrame(tab2, text='信息查询',width=800,heigh=300)
lf2.grid(row=0,column=0, padx=8, pady=4) 
lf2_1 = ttk.LabelFrame(tab2, text='分类汇总',width=800)
lf2_1.grid(row=1,column=0,  padx=8, pady=4)  

label2_1 = ttk.Label(lf2,text='列名：',width=12)
label2_1.grid(row=0,column=0, sticky='W',padx=6,pady=10)
label2_1_stringvar = tk.StringVar()  
label2_1_Chosen = ttk.Combobox(lf2, width=12, textvariable=label2_1_stringvar)  
label2_1_Chosen['values'] = list_label  
label2_1_Chosen.grid(row=0,column=1)  
label2_1_Chosen.current(0)  #设置初始显示值，值为元组['values']的下标  
label2_1_Chosen.config(state='readonly')  #设为只读模式 
label2_1_stringvar2 = tk.StringVar() 
label2_1_entry = tk.Entry(lf2,width=30,textvariable=label2_1_stringvar2)
label2_1_entry.grid(row=0,column=2,sticky='W',padx=6,pady=10)
label2_2 = ttk.Label(lf2,text='列名：')
label2_2.grid(row=1,column=0, sticky='W',padx=6,pady=10)
label2_2_stringvar = tk.StringVar()  
label2_2_Chosen = ttk.Combobox(lf2, width=12, textvariable=label2_2_stringvar)  
label2_2_Chosen['values'] = list_label  
label2_2_Chosen.grid(row=1,column=1)  
label2_2_Chosen.current(0)  #设置初始显示值，值为元组['values']的下标  
label2_2_Chosen.config(state='readonly')  #设为只读模式 
label2_2_stringvar2 = tk.StringVar() 
label2_2_entry = tk.Entry(lf2,width=30,textvariable=label2_2_stringvar2)
label2_2_entry.grid(row=1,column=2,sticky='W',padx=6,pady=10)
label2_3 = ttk.Label(lf2,text='列名：')
label2_3.grid(row=2,column=0, sticky='W',padx=6,pady=10)
label2_3_stringvar = tk.StringVar()  
label2_3_Chosen = ttk.Combobox(lf2, width=12, textvariable=label2_3_stringvar)  
label2_3_Chosen['values'] = list_label  
label2_3_Chosen.grid(row=2,column=1)  
label2_3_Chosen.current(0)  #设置初始显示值，值为元组['values']的下标  
label2_3_Chosen.config(state='readonly')  #设为只读模式 
label2_3_stringvar2 = tk.StringVar() 
label2_3_entry = tk.Entry(lf2,width=30,textvariable=label2_3_stringvar2)
label2_3_entry.grid(row=2,column=2,sticky='W',padx=6,pady=10)
container_tree = tk.Frame(lf2, width=600, height=200)
container_tree.propagate(False)
container_tree.grid(row=3,column=0,columnspan=4,sticky='W',padx=6,pady=10)
    
tree = ttk.Treeview(container_tree,show='headings')
fr_y = tk.Frame(container_tree)
fr_y.pack(side='right', fill='y')
tk.Label(fr_y, borderwidth=1, relief='raised', font="Arial 8").pack(side='bottom', fill='x')
sb_y = tk.Scrollbar(fr_y, orient="vertical", command=tree.yview)
sb_y.pack(expand='yes', fill='y')
fr_x = tk.Frame(container_tree)
fr_x.pack(side='bottom', fill='x')
sb_x = tk.Scrollbar(fr_x, orient="horizontal", command=tree.xview)
sb_x.pack(expand='yes', fill='x')
tree.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
tree.pack(fill='both', expand='yes')
tree["columns"]=list_label
for lab_name in list_label:
    tree.column(lab_name,width=60)   #表示列,不显示  
    tree.heading(lab_name,text=lab_name)  #显示表头  


button_select = tk.Button(lf2,width=12,text="进行筛选",command=select_infor)
button_select.grid(row=0,column=3,padx=8, pady=4) 
button_select = tk.Button(lf2,width=12,text="保存到excel",command=save_to_xls)
button_select.grid(row=1,column=3,padx=8, pady=4) 
 
label2_4 = ttk.Label(lf2_1,text='列名：',width=5)
label2_4.grid(row=0,column=0, sticky='W',padx=6,pady=10)
label2_4_stringvar = tk.StringVar()  
label2_4_Chosen = ttk.Combobox(lf2_1, width=12, textvariable=label2_4_stringvar)  
label2_4_Chosen['values'] = list_label  
label2_4_Chosen.grid(row=0,column=1,sticky='W')  
label2_4_Chosen.current(0)  #设置初始显示值，值为元组['values']的下标  
label2_4_Chosen.config(state='readonly')  #设为只读模式 
button_label2_4 = tk.Button(lf2_1,width=12,text="汇总统计",command=select_infor)
button_label2_4.grid(row=0,column=2,sticky='W',padx=6,pady=10)
scr = scrolledtext.ScrolledText(lf2_1, width=85, height=5, wrap=tk.WORD,state='disabled')  
scr.grid(column=0, row=1, sticky='WE', columnspan=3) 


#select name from syscolumns where id=(select max(id) from sysobjects where xtype='u' and name='persons')
windows.mainloop()