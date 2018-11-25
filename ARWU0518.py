# -*- coding: utf-8 -*-
"""
Created on Sun Nov 25 20:42:43 2018

@author: Cyril Tang
"""

import bs4
import requests
import re
import pandas as pd

host = 'http://www.shanghairanking.com/ARWU2018.html'

def ScrapingRank(host):
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
        kk = symbol.select('td')[3:11]
        for i in kk:
            td.append(i.get_text())
        school.append(td)
        
    res = pd.DataFrame(school)
    res.columns=['World Rank','Institution','Location','Domestic Rank','Total Score','Alumni','Award','HiCi','N&S','PUB','PCP']
    return res


if __name__ == "__main__":
    for year in range(2005,2019):
        host = 'http://www.shanghairanking.com/ARWU'+str(year)+'.html'
        res = ScrapingRank(host)
        res.to_excel('SchoolRank'+str(year)+'.xlsx', sheet_name='Sheet1')