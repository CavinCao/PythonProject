#!/usr/bin/env python3
# -*- coding: utf-8 -*-

__author__ = 'Cavin Cao'

'''
	功能：获取智联招聘网赏的简历
    配置：1.前提你需要有个企业账号
         2.登陆后从网页中获取cookie放至cookies.txt中
         3.维护你想查找的关键词：KEYWORDS
'''

import requests
import json
import config
import pymysql
import time
import random

PAGE=128

KEYWORDS=''

URL='https://rdapi.zhaopin.com/rd/search/resumeList'

HEADER={  
    'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.181 Safari/537.36',
    'zp-route-meta':'uid=668797040,orgid=112956490'
}

def get_cookies():
    f=open("zl_resume_tools/cookies.txt",encoding='utf-8')#zl_jd_tools/
    cookies={}
    for line in f.read().split(';'):
        name,value=line.strip().split('=',1)  
        cookies[name]=value
    return cookies

def get_search_resumeList(index,**cookies):
    data={
        'S_DISCLOSURE_LEVEL':2,
        'S_ENGLISH_RESUME':'1',
        'S_KEYWORD':KEYWORDS,
        'isrepeat':1,
        'rows':30,
        'start':index,
        'sort':'date'
    }
    try:
        res=requests.post(URL,headers=HEADER,data=json.dumps(data),cookies=cookies)
        return res.json()
    except Exception as e:
        print(e)
        return ''

def get_resume_detail(**cookies):
    detail_url='http://ihr.zhaopin.com/resume/details/?access_token=1c42f8eb45884e328ee62710e9489ff9&resumeNo=TAf1mIFj%28skAVlb8rVJr7g_1_1&resumeSource=1&version=3&openFrom=1&haveEnglish=&t=1524978832617&k=5F8650B7A2A8A90D165C7D1DE2D33CD9&searchresume=1&v=1'
    header={
        'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.181 Safari/537.36'
    }

    res=requests.get(detail_url,headers=HEADER,cookies=cookies)
    tree=bs4.BeautifulSoup(res.text,'lxml')
    print(res.text)

def check_repeat_resume(conn,rid):
    try:
        cursor = conn.cursor(cursor=pymysql.cursors.DictCursor)
        sql="SELECT rid FROM base_resume where rid=%s"
        effect_row=cursor.execute(sql,rid)
        return effect_row>0
    except Exception as e:
        return e

def insert_base_resume(conn,rid,rtext):
    cursor = conn.cursor(cursor=pymysql.cursors.DictCursor)
    sql="INSERT INTO base_resume(rid,rtext) values(%s,%s);"
    try:
        cursor.execute(sql,(rid,rtext))
        conn.commit()
        cursor.close()
        return 'success'
    except Exception as e:
        return e
    
def get_resume_main():
    cookies=get_cookies()
    db=config.testdb
    conn=pymysql.connect(host=db['host'],port=db['port'], user=db['user'], passwd=db['password'], db=db['db'],charset='utf8') 
    for num in range(1,PAGE):
        resumeRes=get_search_resumeList(num,**cookies)
        if resumeRes=='':
            continue
        else:
            try:
                resumeList=resumeRes['data']['dataList']
                for resumeModel in resumeList:
                    rid=resumeModel['id']
                    rtext=json.dumps(resumeModel, ensure_ascii=False)
                    isResume=check_repeat_resume(conn,rid)
                    if isResume==False:
                        result=insert_base_resume(conn,rid,rtext)
                        print(rid+':'+str(result))
                    else:
                        print(rid+':已存在')
            except:
                print(str(resumeRes))
        time.sleep(random.randint(3,8))

        

if __name__ =='__main__':
    #cc=get_cookies()
    #get_resume_detail(**cc)
    #get_search_resumeList(1,**cc)
    get_resume_main()
