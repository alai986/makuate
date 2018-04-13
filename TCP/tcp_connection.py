#coding:utf-8
"""
Creater on 2018/4/13
author:alai
"""

from socket import *
import time
import xlwt

# from datetime import date,datetime
#操作Excel


def excel(ldata,ltime):
    # 获取当前时间
    global local_day1,ldata1,ltime1
    local_day = time.strftime('%Y-%m-%d', time.localtime(time.time()))

    # local_time = time.strftime('%H:%M:%S', time.localtime(time.time()))
    name = "马夸特" + local_day

    book = xlwt.Workbook()
    # 新建工作表,可对同一个单元格重复操作
    table = book.add_sheet('塌陷值', cell_overwrite_ok=True)
    # 保存文件,不支持xlsx格式
    # 初始化样式
    style = xlwt.XFStyle()
    # 创建字体
    font = xlwt.Font()
    # 指定字体名字
    font.name = 'Times New Roman'
    # 字体加粗
    font.bold = False
    # 将该font设定为style的字体
    style.font = font

    # 写入到文件时使用该样式
    # data=tcpserver()
    table.write(0, 0, "number", style)
    table.write(0, 1, "塌陷值", style)
    table.write(0, 2, '时间', style)

    if local_day != local_day1:
        local_day1 = local_day
        ldata1 = [ldata]
        ltime1 = [ltime]
    else:
        pass
    for i in range(0, len(ldata)):
        table.write(i + 1, 0, i + 1, style)
        table.write(i + 1, 1, ldata[i], style)
        table.write(i + 1, 2, ltime[i], style)

    # 修改完要再一次保存
    try:
        book.save(name + ".xls")
    except Exception as e:
        print("NG")

def tcpserver():
    global ldata1,ltime1
    host = '127.0.0.1'
    port = 1234
    buffsize = 2048
    tctime = socket(AF_INET, SOCK_STREAM)
    tctime.bind((host,port))
    tctime.listen(3)


    while True:
        print('Wait for connection ...')
        tctimeClient, addr = tctime.accept()
        print("Connection from :", addr)

        while True:
            data = tctimeClient.recv(buffsize).decode()
            if not data:
                break
            tctimeClient.send(('[%s] %s' % (time.ctime(), data)).encode())
            # print("Connection from :", data)

            ldata1.append(data)
            local_time = time.strftime('%H:%M:%S', time.localtime(time.time()))
            # local_time=time.ctime()
            ltime1.append(local_time)
            excel(ldata1,ltime1)


        # tctimeClient.close()
    # tctimeClient.close()
if __name__=="__main__":
    global ldata1
    ldata1 = []
    global ltime1
    ltime1=[]
    global local_day1
    local_day1 = "1970-1-1"

    while True:
        tcpserver()












