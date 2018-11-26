import time              
import re                 
from selenium import webdriver                  
import selenium.webdriver.support.ui as ui           
import xlwt
import xlrd 
import os 

#driver = webdriver.Firefox()
#wait = ui.WebDriverWait(driver,10)  

inforead = xlrd.open_workbook(r'IDs.xlsx') 
sheet1 = inforead.sheet_by_index(0) 

def LoginWeibo(username, password):  
    try:  
        #**********************************************************************  
        # 直接访问driver.get("http://weibo.cn/5824697471")会跳转到登陆页面 用户id  
        #  
        # 用户名<input name="mobile" size="30" value="" type="text"></input>  
        # 密码 "password_4903" 中数字会变动,故采用绝对路径方法,否则不能定位到元素  
        #  
        # 勾选记住登录状态check默认是保留 故注释掉该代码 不保留Cookie 则'expiry'=None  
        #**********************************************************************  
          
        #输入用户名/密码登录  
        print (u'准备登陆Weibo.cn网站...')  
        driver.get("http://login.weibo.cn/login/")  
        elem_user = driver.find_element_by_name("mobile")  
        elem_user.send_keys(username) #用户名  
        elem_pwd = driver.find_element_by_xpath("/html/body/div[2]/form/div/input[2]")  
        elem_pwd.send_keys(password)  #密码  

        #elem_rem = driver.find_element_by_name("remember")  
        #elem_rem.click()             #记住登录状态  


        time.sleep(20)  
          
        elem_sub = driver.find_element_by_name("submit")  
        elem_sub.click()              #点击登陆  
        time.sleep(2)  
          

        for cookie in driver.get_cookies():   
            #print cookie  
            for key in cookie:  
                print (key, cookie[key])  
                      
        driver.get_cookies()#类型list 仅包含一个元素cookie类型dict  
        print (u'登陆成功...')  
          
          
    except (Exception)as e:        
        print ("Error: ",e)  
    finally:      
        print (u'End LoginWeibo!\n\n') 
        
        
def makefolder(user):
    k=0
    while k<=len(user)-1:
        fileposition=r'c:/Ashley/matching/' + str(user[k]) +r'/'
        os.mkdir(fileposition)
        k+=1
    else:
        print('new folder made')

        
        
def VisitPersonPage(user_id):  
    try:
        k=0
        while k<=len(user)-1:  #ID的数目减去一
            f=xlwt.Workbook()
            sheet1 = f.add_sheet(u'sheet1',cell_overwrite_ok=True)
            row0 = [u'id',u'name',u'blogs',u'focus',u'fans',u'origin',u'agree',u'reblog',u'comments',u'time',u'contents']
            for i in range(0,len(row0)):
                sheet1.write(0,i,row0[i])
                
            driver.get("http://weibo.cn/" + str(user_id[k]))
            m=1
            str_name = driver.find_element_by_xpath("//div[@class='ut']")  
            str_t = str_name.text.split(" ")  
            num_name = str_t[0]      #空格分隔 获取第一个值 "Eastmount 详细资料 设置 新手区"  
            print (u'昵称: ' + num_name)
                       
            str_wb = driver.find_element_by_xpath("//div[@class='tip2']")    
            pattern = r"\d+\.?\d*"   #正则提取"微博[0]" 但r"(.∗?)"总含[]   
            guid = re.findall(pattern, str_wb.text, re.S|re.M)  
            print (str_wb.text)        #微博[294] 关注[351] 粉丝[294] 分组[1] @他的  
            for value in guid:  
                num_wb = int(value)
                break  
            print (u'微博数: ' + str(num_wb))
                
            str_gz = driver.find_element_by_xpath("//div[@class='tip2']/a[1]")  
            guid = re.findall(pattern, str_gz.text, re.M)  
            num_gz = int(guid[0])  
            print (u'关注数: ' + str(num_gz))
                
            str_fs = driver.find_element_by_xpath("//div[@class='tip2']/a[2]")  
            guid = re.findall(pattern, str_fs.text, re.M)  
            num_fs = int(guid[0])  
            print (u'粉丝数: ' + str(num_fs))
            
            n=1 
            while n<=8:   #要跑几页微博，每个页面有十条微博信息
                url_wb = "http://weibo.cn/" + user_id[k] + "?filter=0&page=" + str(n)
                driver.get(url_wb)
                info = driver.find_elements_by_xpath("//div[@class='c']")
                for value in info:    
                    info = value.text
                    
                    if u'设置:皮肤.图片' not in info:
                        
                        str4 = info.split(u" 收藏 ")[-1]  
                        flag = str4.find(u"来自")  
                        print (u'时间: ' + str4[:flag])
                        time = str4[:flag]
                        if u'04月14日' in time:
                                
    
                            if info.startswith(u'转发'):  
                                print (u'转发微博')
                                ori = 0
                            else:  
                                print (u'原创微博')
                                ori = 1
                        #获取最后一个点赞数 因为转发是后有个点赞数  
                            str1 = info.split(u" 赞")[-1]  
                            if str1:   
                                val1 = re.match(r"\[([^\[\]]*)\]", str1).groups()[0]  
                                print (u'点赞数: ' + val1)
          
                            str2 = info.split(u" 转发")[-1]  
                            if str2:   
                                val2 = re.match(r'\[([^\[\]]*)\]', str2).groups()[0]  
                                print (u'转发数: ' + val2)
          
                            str3 = info.split(u" 评论")[-1]  
                            if str3:  
                                val3 = re.match(r'\[([^\[\]]*)\]', str3).groups()[0]  
                                print (u'评论数: ' + val3)
          
          
                            print (u'微博内容:')  
                            print (info[:info.rindex(u" 赞")])  #后去最后一个赞位置
                            val4 = str(info[:info.rindex(u" 赞")])
                            print ('\n' )
                        
                            row1 = [user_id[k],num_name,num_wb,num_gz,num_fs,ori,val1,val2,val3,time,val4]
                            for i in range(0,len(row1)):
                                sheet1.write(m,i,row1[i])
    
                            m+=1
                        else:
                            print('old blog')
                    else:  
                        print (u'跳过', info, '\n')  
                        break
                    
                else:
                    print('next page')
                n+=1
            
            fileposition=r'c:/Ashley/blogs/' + str(user[k]) +r'/'
            filename = str(user[k])+'0414'+'.xls'
            k+=1
            f.save(fileposition + filename )

    except (Exception)as e:        
                print ("Error: ",e)  
    finally:
        print("End")
        
if __name__ == '__main__':  
  
    #定义变量  
    username = '18805927816'             #输入你的用户名  
    password = '5love5jia'               #输入你的密码  
  
  
    #操作函数  
    #LoginWeibo(username, password)      #登陆微博  
  
    #driver.add_cookie({'name':'name', 'value':'_T_WM'})  
    #driver.add_cookie({'name':'value', 'value':'c86fbdcd26505c256a1504b9273df8ba'})  
  
    #注意  
    #因为sina微博增加了验证码,但是你用Firefox登陆一次输入验证码,再调用该程序即可,因为Cookies已经保证  
    #会直接跳转到明星微博那部分,即: http://weibo.cn/guangxianliuyan  
      
  
    #在if __name__ == '__main__':引用全局变量不需要定义 global inforead 省略即可  
    print ('Read file:')  
    user_id = sheet1.col_values(0)
    user = sheet1.col_values(1)
    makefolder(user)
    #VisitPersonPage(user_id)            

            #访问个人页面    
  