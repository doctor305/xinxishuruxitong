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
import re
from connect_to_mssql import *
from write_xls import *


def get_server_info():
    server_config = []
    with open('config.ini','r') as f:
        line = f.readline()
        while line:
            server_config.append(line.split('=')[1].strip())
            line = f.readline()
    return server_config


def input_infor():
    global conn,display_vartext
    list_value = []
    if list_entry[3].get() == '':
        mBox.showwarning("错误提示","身份证号不能为空")
    else:
        for e in list_entry:
            if e.get() != '':
                list_value.append(e.get())
            else:
                list_value.append('Null')
            e.set('')
        insert_value(conn,list_value)
        display_vartext.set('输入信息已保存！')
      

def select_infor():
    global conn
    condition1 = "%s like '%%%s%%'" % (label2_1_stringvar.get(),label2_1_stringvar2.get())
    condition2 = "%s like '%%%s%%'" % (label2_2_stringvar.get(),label2_2_stringvar2.get())
    condition3 = "%s like '%%%s%%'" % (label2_3_stringvar.get(),label2_3_stringvar2.get())
    condition = condition1 + ' and ' + condition2 + ' and ' + condition3
    if chVarUn_age.get() == 0:
        commit = 'select * from informationtable where %s' % condition
    else:
        if not checkdate1() or not checkdate2():
            mBox.showwarning("格式不规范","格式不规范，出生年月请填入4位的数字！")
            return
        else:
            if entryage_1_stringvar1.get()== '' and entryage_2_stringvar1.get()== '':
                commit = 'select * from informationtable where %s' % condition
            elif entryage_1_stringvar1.get()== '':
                commit = u'select * from informationtable where %s and 出生年月 <= %d' % (condition,int(entryage_2_stringvar1.get()))
            elif entryage_2_stringvar1.get()== '':
                commit = u'select * from informationtable where %s and 出生年月 >= %d' % (condition,int(entryage_1_stringvar1.get()))
            else:
                commit = u'select * from informationtable where %s and 出生年月>=%d and 出生年月 <=%d' % (condition,int(entryage_1_stringvar1.get()),int(entryage_2_stringvar1.get()))
    result = send_select(conn,commit.encode('utf8'))
    items = tree.get_children()
    for item in items:
        tree.delete(item) #删除原有数据
    for i in range(len(result)):
        tree.insert("",i,values=result[i]) #插入数据
        
def save_to_xls():
    values = []
    filename=asksaveasfilename(parent = windows,defaultextension='.xls')
    if filename != '':
        items = tree.get_children()
        for item in items:
            values.append(tree.item(item,'values'))
        write_xls(list_label,values,filename)

def checkdate_4(string):
    if re.search(r'^[0-9]{4}$',string):
        return True
    else:
        return False
def checkdate_8(string):
    if re.search(r'^[0-9]{8}$',string):
        return True
    else:
        return False
def checkdate1():
    if checkdate_4(entryage_1_stringvar1.get()) or entryage_1_stringvar1.get()=='':
        return True
    else:
        return False
def checkdate2():
    if checkdate_4(entryage_2_stringvar1.get()) or entryage_2_stringvar1.get()=='':
        return True
    else:
        return False

def msg_datecheck():
    mBox.showwarning("格式不规范","格式不规范，请填入4位的数字！")
    
def select_group():
    commit = u"select %s,count(1) 人数 from informationtable group by %s" % (label2_4_stringvar.get(),label2_4_stringvar.get())
    result = send_select(conn,commit.encode('utf8'))
    scr.delete('1.0',tk.END)
    for i in result:
        string = u'%s %s 共有 %s 人\n' % (label2_4_stringvar.get(),i[0],i[1])
        scr.insert(tk.END, string)
    scr.see(tk.END)

def search_infor():
    global search_sfzh
    if vartext_search.get()=='':
        mBox.showwarning("缺少信息","请填入搜索所需的身份证号！")
        button3_2.configure(state='disabled')
        button3_3.configure(state='disabled')
    else:
        search_sfzh = vartext_search.get().strip()
        commit = u"select * from informationtable where 身份证号 like '%s'" % search_sfzh
        result = send_select(conn,commit.encode('utf8'))
        if len(result)==0:
            mBox.showwarning("没有相符信息","未找到符合此身份证号的信息，请重新输入！")
            button3_2.configure(state='disabled')
            button3_3.configure(state='disabled')
        else:
            for i in range(len(list_label)):
                list_entry3[i].set(result[0][i])
            button3_2.configure(state='normal')
            button3_3.configure(state='normal')

def update_infor():
    global search_sfzh
    commit_element = []
    for i in range(len(list_label)):
        commit_element.append(list_label[i]+"='"+list_entry3[i].get()+"'")
    commit_temp = ",".join(commit_element)
    commit = u"update informationtable set %s where 身份证号 like '%s'" % (commit_temp,search_sfzh)
    answer = mBox.askyesno(u"确认", u"是否确定对此条信息进行更改？")   
    if answer == True:
        send_commit(conn,commit.encode('utf8'))
        button3_2.configure(state='disabled')
        button3_3.configure(state='disabled')
        for j in list_entry3:
            j.set('')
        mBox.showinfo(u'修改信息', u'该条信息已修改！')  
    
host,user,password,database = get_server_info()

conn = connect_to_mssql(host,user,password,database)
check_table(conn)
search_sfzh = ''

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

label2_1 = ttk.Label(lf2,text='列名：',width=5)
label2_1.grid(row=0,column=0, sticky='W',padx=6,pady=10)
label2_1_stringvar = tk.StringVar()  
label2_1_Chosen = ttk.Combobox(lf2, width=8, textvariable=label2_1_stringvar)  
label2_1_Chosen['values'] = list_label  
label2_1_Chosen.grid(row=0,column=1)  
label2_1_Chosen.current(0)  #设置初始显示值，值为元组['values']的下标  
label2_1_Chosen.config(state='readonly')  #设为只读模式 
label2_1_stringvar2 = tk.StringVar() 
label2_1_entry = tk.Entry(lf2,width=16,textvariable=label2_1_stringvar2)
label2_1_entry.grid(row=0,column=2,sticky='W',padx=6,pady=10)
label2_2 = ttk.Label(lf2,text='列名：',width=5)
label2_2.grid(row=1,column=0, sticky='W',padx=6,pady=10)
label2_2_stringvar = tk.StringVar()  
label2_2_Chosen = ttk.Combobox(lf2, width=8, textvariable=label2_2_stringvar)  
label2_2_Chosen['values'] = list_label  
label2_2_Chosen.grid(row=1,column=1)  
label2_2_Chosen.current(0)  #设置初始显示值，值为元组['values']的下标  
label2_2_Chosen.config(state='readonly')  #设为只读模式 
label2_2_stringvar2 = tk.StringVar() 
label2_2_entry = tk.Entry(lf2,width=16,textvariable=label2_2_stringvar2)
label2_2_entry.grid(row=1,column=2,sticky='W',padx=6,pady=10)
label2_3 = ttk.Label(lf2,text='列名：',width=5)
label2_3.grid(row=2,column=0, sticky='W',padx=6,pady=10)
label2_3_stringvar = tk.StringVar()  
label2_3_Chosen = ttk.Combobox(lf2, width=8, textvariable=label2_3_stringvar)  
label2_3_Chosen['values'] = list_label  
label2_3_Chosen.grid(row=2,column=1)  
label2_3_Chosen.current(0)  #设置初始显示值，值为元组['values']的下标  
label2_3_Chosen.config(state='readonly')  #设为只读模式 
label2_3_stringvar2 = tk.StringVar() 
label2_3_entry = tk.Entry(lf2,width=16,textvariable=label2_3_stringvar2)
label2_3_entry.grid(row=2,column=2,sticky='W',padx=6,pady=10)
container_tree = tk.Frame(lf2, width=600, height=200)
container_tree.propagate(False)
container_tree.grid(row=3,column=0,columnspan=5,sticky='W',padx=6,pady=10)
    
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

lf_age = ttk.LabelFrame(lf2, text='年龄筛选',width=50,heigh=100)
lf_age.grid(row=0, column=3,rowspan=3,sticky='W' )
chVarUn_age = tk.IntVar()  
check_age = tk.Checkbutton(lf_age, text="启用", variable=chVarUn_age)  
check_age.deselect()   #清除 (turns off) checkbutton.  
check_age.grid(row=0, column=0,sticky='W' )                    

labelage_1 = ttk.Label(lf_age,text='出生晚于',width=8)
labelage_1.grid(row=1,column=0, sticky='W',pady=10)
entryage_1_stringvar1 = tk.StringVar() 
entryage_1 = tk.Entry(lf_age,width=10,textvariable=entryage_1_stringvar1,\
                      validate='focusout',validatecommand=checkdate1,\
                      invalidcommand=msg_datecheck)
entryage_1.grid(row=1,column=1, sticky='W',pady=10)
labelage_1_1 = ttk.Label(lf_age,text='年',width=2)
labelage_1_1.grid(row=1,column=2, sticky='W',pady=10)

labelage_2 = ttk.Label(lf_age,text='出生早于',width=8)
labelage_2.grid(row=2,column=0, sticky='W',pady=10)
entryage_2_stringvar1 = tk.StringVar() 
entryage_2 = tk.Entry(lf_age,width=10,textvariable=entryage_2_stringvar1,\
                      validate='focusout',validatecommand=checkdate2,\
                      invalidcommand=msg_datecheck)
entryage_2.grid(row=2,column=1, sticky='W',pady=10)
labelage_2_1 = ttk.Label(lf_age,text='年',width=2)
labelage_2_1.grid(row=2,column=2, sticky='W',pady=10)


button_select = tk.Button(lf2,width=12,text="进行筛选",command=select_infor)
button_select.grid(row=0,column=4,padx=8, pady=4) 
button_select = tk.Button(lf2,width=12,text="保存到excel",command=save_to_xls)
button_select.grid(row=1,column=4,padx=8, pady=4) 
 
label2_4 = ttk.Label(lf2_1,text='分组所用列名：',width=15)
label2_4.grid(row=0,column=0, sticky='W',padx=6,pady=10)
label2_4_stringvar = tk.StringVar()  
label2_4_Chosen = ttk.Combobox(lf2_1, width=12, textvariable=label2_4_stringvar)  
label2_4_Chosen['values'] = list_label  
label2_4_Chosen.grid(row=0,column=1,sticky='W')  
label2_4_Chosen.current(0)  #设置初始显示值，值为元组['values']的下标  
label2_4_Chosen.config(state='readonly')  #设为只读模式 
button_label2_4 = tk.Button(lf2_1,width=12,text="汇总统计",command=select_group)
button_label2_4.grid(row=0,column=2,sticky='W',padx=6,pady=10)

scr = scrolledtext.ScrolledText(lf2_1, width=85, height=5, wrap=tk.WORD)  
scr.grid(column=0, row=1, sticky='WE', columnspan=3) 

# Notebook 第三页面控件
lf3 = ttk.LabelFrame(tab3, text='信息修改')
lf3.grid(column=0, row=0, padx=8, pady=4)  

n = 0
list_entry3 = []
if len(list_label)%2 == 0:
    tag_number = len(list_label)//2
else:
    tag_number = len(list_label)//2+1
while n<len(list_label):
    vartext = tk.StringVar()
    list_entry3.append(vartext)
    if n < tag_number:
        label = ttk.Label(lf3,text=list_label[n])
        label.grid(row=n,column=0, padx=6,pady=10, sticky=tk.W)
        entry = tk.Entry(lf3,width=30,textvariable=list_entry3[n])
        entry.grid(row=n,column=1,padx=6,pady=10, sticky=tk.W)
    else:
        label = ttk.Label(lf3,text=list_label[n])
        label.grid(row=n-tag_number,column=2,padx=6,pady=10, sticky=tk.W)
        entry = tk.Entry(lf3,width=30,textvariable=list_entry3[n])
        entry.grid(row=n-tag_number,column=3,padx=6,pady=10, sticky=tk.W)
    n += 1
lf3_2 = ttk.LabelFrame(lf3, text='查找')
lf3_2.grid(row=n+1,column=0,columnspan=4,padx=8, pady=4)

label_search = ttk.Label(lf3_2,text=u'按身份证号查找')
label_search.grid(row=0,column=0,padx=6,pady=10, sticky=tk.W)
vartext_search = tk.StringVar()
entry_search = tk.Entry(lf3_2,width=20,textvariable=vartext_search)
entry_search.grid(row=0,column=1,padx=6,pady=10, sticky=tk.W)
button3_1 = tk.Button(lf3_2,width=10,text="查找",command=search_infor)
button3_1.grid(row=0,column=2, padx=8, pady=4)    
button3_2 = tk.Button(lf3_2,width=10,text="修改",state='disable',command=update_infor)
button3_2.grid(row=0,column=3, padx=8, pady=4)
button3_3 = tk.Button(lf3_2,width=10,text="删除",state='disable',command=update_infor)
button3_3.grid(row=0,column=4, padx=8, pady=4)



##          添加菜单
def _quit():  
    windows.quit()  
    windows.destroy()  
    exit()  
      
menuBar = Menu(windows)  
windows.config(menu=menuBar)  
  
fileMenu = Menu(menuBar, tearoff=0)  
fileMenu.add_command(label="新建")  
fileMenu.add_separator()  
fileMenu.add_command(label="退出", command=_quit)  
menuBar.add_cascade(label="文件", menu=fileMenu) 

#select name from syscolumns where id=(select max(id) from sysobjects where xtype='u' and name='persons')
windows.mainloop()
