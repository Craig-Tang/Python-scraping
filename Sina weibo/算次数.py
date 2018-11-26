# -*- coding: utf-8 -*-
"""
Created on Sat Apr 22 14:52:39 2017

@author: Cyril Tang
"""

import xlrd
#import xlwt
from xlutils.copy import copy
from collections import Counter

lis=xlrd.open_workbook('c:/Ashley/python/IDs.xlsx')
sheet=lis.sheet_by_index(0)

li=[]

def match(user,time):
    hotfile='C:/Ashley/hotall/' + time + '.xls'
    hot=xlrd.open_workbook(hotfile)
    s1=hot.sheet_by_index(0)
    n1=s1.nrows

    for l in range(0,len(user)-1):
        blogfile='c:/Ashley/matching/'+user[l]+'/'+user[l]+time+'matched.xls'
        con=xlrd.open_workbook(blogfile)
        s2=con.sheet_by_index(0)
        n2=s2.nrows
        '''
        ma=copy(con)
        ma.get_sheet(0).write(0,12,'num')
        ma.get_sheet(0).write(0,11,'match')
        '''
        #for i in range(1,n2):
            #ma.get_sheet(0).write(i,11,0)

        
        j=1
        while j <= n1-1:
            key=s1.row(j)[0].value
            i=1
            while i<=n2-1:
                print(i)
                text=s2.row(i)[10].value
                k=0
                while k <= len(key)-1:
                    w=text.find(key[k])
                    #ma.get_sheet(0).write(i,12,l+1)
                    if w != -1:
                        k+=1
                        if k==len(key):
                            print('find it!')
                            #ma.get_sheet(0).write(i,11,1)
                            li.append(key)
                    else:
                        break
                i+=1
            j+=1
        #blognew='c:/Ashley/matching/'+ user[l]+'/'+user[l]+time+'matched.xls'
        #ma.save(blognew)
        
if __name__ == '__main__':
    user=sheet.col_values(1)
    time='0410'
    match(user,time)
    print(Counter(li).most_common())
