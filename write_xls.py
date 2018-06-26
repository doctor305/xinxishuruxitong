#!/usr/bin/evn python
#-*- coding:utf-8 -*-
'''
Created on 2018年6月26日

@author: jinfeng
'''
import xlwt

def write_xls(column_name,values,save_path): 
    book = xlwt.Workbook(encoding = 'utf8',style_compression=0)
    sheet = book.add_sheet('查询结果')
    for n in range(len(column_name)):
        sheet.write(0,n,column_name[n])
    for m in range(len(values)):
        for l in range(len(column_name)):
            sheet.write(m+1,l,values[m][l])
    book.save(save_path)
        
   
if __name__ == '__main__':
    write_xls()