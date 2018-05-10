#!/usr/bin/env python3
# -*- coding: utf-8 -*-

__author__ = 'Cavin Cao'

'''
	功能：落库的数据解析导出excel
'''

import xlsxwriter
import requests
import json
import config
import pymysql
import re

def get_resume_list():
    db=config.testdb
    conn=pymysql.connect(host=db['host'],port=db['port'], user=db['user'], passwd=db['password'], db=db['db'],charset='utf8') 
    try:
        cursor = conn.cursor(cursor=pymysql.cursors.DictCursor)
        sql="SELECT * FROM base_resume "
        effect_row=cursor.execute(sql)
        result=  cursor.fetchall()
        cursor.close()
        conn.close()
        return result
    except Exception as e:
        print('读取异常：'+e)

def etl_resume_main():
    workbook = xlsxwriter.Workbook('resume_list.xlsx')
    worksheet = workbook.add_worksheet('风资源')
    worksheet.write(0,0, '简历ID') 
    worksheet.write(0,1, '简历别名') 
    worksheet.write(0,2, '姓名') 
    worksheet.write(0,3, '职位') 
    worksheet.write(0,4, '学历') 
    worksheet.write(0,5, '性别')
    worksheet.write(0,6, '年龄')
    worksheet.write(0,7, '现居住地')
    worksheet.write(0,8, '期望月薪')
    worksheet.write(0,9, '当前状态')
    worksheet.write(0,10, '工龄')
    worksheet.write(0,11, '学校')
    worksheet.write(0,12, '专业')
    worksheet.write(0,13, '期望工作地点')
    worksheet.write(0,14, '公司')

    resumeList=get_resume_list()
    i=1
    for resume in resumeList:
        #resumeDetail=json.loads(resume['rtext'])
        try:
            regex = re.compile(r'\\(?![/u"])')  
            fixed = regex.sub(r"\\\\", resume['rtext']) 
            resumeDetail=json.loads(fixed, strict=False)  
            worksheet.write(i,0, resume['rid'])
            worksheet.write(i,1, resumeDetail['name'])
            worksheet.write(i,2, resumeDetail['userName'])
            worksheet.write(i,3, resumeDetail['jobTitle'])
            worksheet.write(i,4, resumeDetail['eduLevel'])
            worksheet.write(i,5, resumeDetail['gender'])
            worksheet.write(i,6, resumeDetail['age'])
            worksheet.write(i,7, resumeDetail['city'])
            worksheet.write(i,8, resumeDetail['desiredSalary'])
            worksheet.write(i,9, resumeDetail['careerStatus'])
            worksheet.write(i,10, resumeDetail['workYears'])
            worksheet.write(i,11, resumeDetail['school'])
            worksheet.write(i,12, resumeDetail['major'])
            worksheet.write(i,13, resumeDetail['desireCity'])
            worksheet.write(i,14, resumeDetail.get('lastJobDetail','').get('companyName',''))
        except Exception as e:
            print(e)
            worksheet.write(i,0, resume['rid'])
        i=i+1

    workbook.close()

if __name__ =='__main__':
    etl_resume_main()