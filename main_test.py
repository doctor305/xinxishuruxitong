#!/usr/bin/evn python
#-*- coding:utf-8 -*-
'''
Created on 2018年6月19日

@author: jinfeng
'''

import os
import pymssql
import collections

def creat_configfile():
    global host,port,user,password
    host = '127.0.0.1'
    port = '1433'
    user = 'sa'
    password = '12345678'
    with open('config.ini','w') as f:
        f.write(host+':'+port+':'+user+':'+password)

if os.path.exists('config.ini'):
    try:
        with open('config.ini','r') as f:
            host,port,user,password = f.read().split(":")
            print(host,port,user,password)
    except:
        creat_configfile()
        print(host,port,user,password)
else:
    creat_configfile()
    print(host,port,user,password)
#conn = pymssql.connect(host, user, password, 'master')
#cursor = conn.cursor()
#cursor.execute("""CREATE DATABASE my_db""")
#conn.commit()
    
conn = pymssql.connect(host, user, password, 'test')
cursor = conn.cursor()
cursor.execute("""
IF OBJECT_ID('persons', 'U') IS NOT NULL
    DROP TABLE persons
CREATE TABLE persons (
    id INT NOT NULL,
    name VARCHAR(100),
    salesrep VARCHAR(100),
    PRIMARY KEY(id)
)
""")
cursor.executemany(
    "INSERT INTO persons VALUES (%d, %s, %s)",
    [(1, 'John Smith', 'John Doe'),
     (2, 'Jane Doe', 'Joe Dog'),
     (3, 'Mike T.', 'Sarah H.')])
# 如果没有指定autocommit属性为True的话就需要调用commit()方法
conn.commit()

#cursor.execute("CREATE DATABASE my_db")
#cursor.execute("select * from 信息表")
#row = cursor.fetchone()
#while row:
#    print(len(row))
#    print('-'.join(row))
#    row = cursor.fetchone()

#conn.commit()
    


    
