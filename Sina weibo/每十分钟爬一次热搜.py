# -*- coding: utf-8 -*-
"""
Created on Fri Mar 24 19:21:01 2017

@author: jilig
"""
import requests  
import re  
import xlwt  
import time  
from bs4 import BeautifulSoup  
from datetime import  datetime, timedelta  

now = datetime.now()  
strnow = now.strftime('%Y-%m-%d %H:%M') 
  
def work():  
    myfile=xlwt.Workbook()  
    table1=myfile.add_sheet(u"实时热搜榜",cell_overwrite_ok=True)  
    table1.write(0,0,u"热搜关键词")  
    table1.write(0,1,u"热搜指数")  

    user_agent = 'Mozilla/4.0 (compatible; MSIE 5.5; Windows NT)'  
    headers = { 'User-Agent' : user_agent }  
#print (soup.prettify())  
    class sousuo():  
        def __init__(self,url,table):  
            self.url=url  
            self.table=table  
          
        def chaxun(self):  
            url = self.url  
            r=requests.get(url,headers=headers)  
            html=r.text  
            
            soup=BeautifulSoup(html)  
            #print (soup.prettify())  
            #获取热搜名称  
            i=1  
            for tag in soup.find_all(href=re.compile("Refer=top"),target="_blank"):  
                if tag.string is not None:  
                    print (tag.string)  
                    self.table.write(i,0,tag.string)  
                    i+=1  
  
            #获取热搜关注数  
            j=1  
            for tag in soup.find_all(class_="star_num"):  
                if tag.string is not None:  
                    print (tag.string)  
                    self.table.write(j,1,tag.string)  
                    j+=1  
  
    s1=sousuo('http://s.weibo.com/top/summary?cate=realtimehot',table1)  
    s1.chaxun() 

    filename=str(time.strftime('%Y-%m-%d_%H%M',time.localtime()))+"热搜.xls"  
    myfile.save(filename)  
    print (u"完成%s的微博热搜备份"%time.strftime('%Y-%m-%d_%H%M',time.localtime()))   

  
def runTask(func, day=0, hour=0, min=0, second=0):  
  # Init time  
 
  print ("now:",strnow ) 
  # First next run time  
  period = timedelta(days=day, hours=hour, minutes=min, seconds=second)  
  next_time = now + period  
  strnext_time = next_time.strftime('%Y-%m-%d %H:%M')  
  print ("next run:",strnext_time ) 
  while True:  
      # Get system current time  
      iter_now = datetime.now()  
      iter_now_time = iter_now.strftime('%Y-%m-%d %H:%M')  
      if str(iter_now_time) == str(strnext_time):  
          # Get every start work time  
          print ("start work: %s" % iter_now_time  )
          # Call task func  
          func()  
          print ("task done.")  
          # Get next iteration time  
          iter_time = iter_now + period  
          strnext_time = iter_time.strftime('%Y-%m-%d %H:%M')  
          print ("next_iter: %s" % strnext_time)  
          # Continue next iteration  
          continue  
  
# runTask(work, min=0.5)  
runTask(work, min=10) 

