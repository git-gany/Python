import pymysql
import time
import xlwt
import MY_CONFIG  # 配置文件，数据库连接相关参数包含其中

def Run():
    
    # 数据库操作
    db = pymysql.connect(address, username, password, data_base)
    try:
        daochu = db.cursor()
        daochu.execute(sql_captcha) # 执行sql
        output_daochu = list(daochu.fetchall()) # 打印结果为tuple类型，将之转化为可插入的list类型
    except Exception as error:
        print('操作数据库时发生异常，请检查网络状况及sql语句的正确性！')
        db.close()
    db.close()

    # 文件操作
    try:
        file = xlwt.Workbook()
        sheet = file.add_sheet('数据',cell_overwrite_ok=True)    
        output_daochu.insert(0,title) # 插入列名
        for i,row in enumerate(output_daochu):
            for j,col in enumerate(row):
                sheet.write(i,j,str(col)) # 坐标及内容
    except Exception as error:
        print('操作文件时发生异常,请检查文件是否处于打开状态！')
        return
    file.save(filename)


#-----导出信息配置------------------------------------------------------------------------------------

address = MY_CONFIG.Data_base_address()
username = MY_CONFIG.Data_base_username()
password = MY_CONFIG.Data_base_password()
data_base = MY_CONFIG.Data_base_name()

# sql语句
sql_captcha = 'SELECT \
                    code,mobile,create_date \
                FROM \
                    t_sms \
                ORDER BY create_date DESC \
                LIMIT 10'

# excel文件列名   list & tuple
title = ['验证码','手机号','发送时间']

# 文件名称
filename = '验证码.xls'

#-----运行-------------------------------------------------------------------------------------

Run()
