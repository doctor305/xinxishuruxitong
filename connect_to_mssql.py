#!/usr/bin/evn python
#-*- coding:utf-8 -*-
'''
Created on 2018年6月21日

@author: jinfeng
'''
import pymssql

def connect_to_mssql(host,user,password,db):
    try:
        conn = pymssql.connect(host, user, password, db)
        return conn
    except:
        return 0

def send_commit(conn,commit):
    cursor = conn.cursor()
    cursor.execute(commit)
    conn.commit()
    
def create_table(conn):
    commit = """
    IF OBJECT_ID('informationtable', 'U') IS NOT NULL
    DROP TABLE informationtable
    CREATE TABLE informationtable (
    姓名 NVARCHAR(255) NOT NULL,
    性别 NVARCHAR(255),
    出生年月 NVARCHAR(255),
    身份证号 NVARCHAR(255) NOT NULL,
    组别 NVARCHAR(255),
    农保 NVARCHAR(255),
    社保 NVARCHAR(255),
    医保 NVARCHAR(255),
    学历 NVARCHAR(255),
    政治面貌 NVARCHAR(255),
    工作单位 NVARCHAR(255),
    联系方式 NVARCHAR(255),
    备注 NVARCHAR(255)
)
"""
    send_commit(conn,commit)

def send_select(conn,commit):
    result = []
    cursor = conn.cursor()
    cursor.execute(commit)
    row = cursor.fetchone()
    while row:
        result.append(row)
        row = cursor.fetchone()
    return result

def get_rowname(conn):
    commit = "select name from syscolumns where id=(select max(id) from sysobjects where xtype='u' and name='informationtable')"
    result = []
    cursor = conn.cursor()
    cursor.execute(commit)
    row = cursor.fetchone()
    while row:
        result.append(row[0])
        row = cursor.fetchone()
    return result

def insert_value(conn,ls):
    commit = "INSERT INTO informationtable VALUES (%s)" % ("'"+"','".join(ls)+"'").replace("'Null'","Null")
    send_commit(conn,commit)

if __name__ == "__main__":
    host = '127.0.0.1'
    user = 'sa'
    password = '12345678'
    database = 'test'
    conn = connect_to_mssql(host,user,password,database)
    #create_table(conn)
    list_rowname = get_rowname(conn)
    print(list_rowname)
    ls = ['张三','男','19920406','410183199204060055','2','Null','Null','Null','大学','党员','Null','Null','Null']
    commit = "'"+"','".join(ls)+"'"
    print(commit)
    insert_value(conn,ls)
