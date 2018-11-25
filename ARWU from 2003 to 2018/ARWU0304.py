# -*- coding: utf-8 -*-
"""
Created on Mon Nov 26 00:03:41 2018

@author:  Cyril Tang
"""

import bs4
import requests
import re
import pandas as pd

#host = 'http://www.shanghairanking.com/ARWU2004.html'
host = 'http://www.shanghairanking.com/ARWU2003.html'

response = requests.get(host)
soup = bs4.BeautifulSoup(response.text,'lxml')

ranklist = soup.select('table')[0].select('tr')[1:]

school = []

for i, symbol in enumerate(ranklist):
    td = []
    td.append(symbol.select('td')[0].get_text())
    td.append(symbol.select('td')[1].get_text())
    td_country = symbol.select('img')[0].get('src')
    td.append(re.search(r'(?<=flag/).*?(?=.png)',td_country).group())
    #kk = symbol.select('td')[3:10]
    kk = symbol.select('td')[3:9]
    for i in kk:
        td.append(i.get_text())
    school.append(td)
    
res = pd.DataFrame(school)
#res.columns=['World Rank','Institution','Location','Total Score','Alumni','Award','HiCi','N&S','PUB','PCP']
res.columns=['World Rank','Institution','Location','Total Score','Nobel','HiCi','N&S','PUB','Faculty']
#res.to_excel('SchoolRank2004.xlsx', sheet_name='Sheet1')
res.to_excel('SchoolRank2003.xlsx', sheet_name='Sheet1')