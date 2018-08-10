# -*- coding: utf-8 -*-
"""
# coding: utf-8 is used to make Chinese encoding compatitable
Created on Fri Aug 10 17:23:24 2018

@author: Phoenix_TAN
"""

import re
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import xlwt
import xlrd
from xlutils.copy import copy
import sys
import time
import datetime
import random

# Record current time
time1 = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) 
print(time1)

# get the source code of the website
option = webdriver.ChromeOptions()  # Set the option
option.add_argument('headless')     # Headless browser
browser = webdriver.Chrome(chrome_options=option) # Add the parament
browser.get('https://movie.douban.com/tag/#/')
# print(browser.page_source)
time.sleep(2)
# Sleep for two seconds until the finish of the initilization of th website
# Because it needs response time and load time to show all the source code

# click the button "Load More"
click_time = 260  # You can set the times of click
counter0 = click_time
while( counter0 ):
    try:
        print(len(browser.page_source))  ## test the length of the source code 
        print("You has clicked for "+str(click_time-counter0)+ " times! ")                 
        browser.find_element_by_class_name("more").click()
        time.sleep(2)   
    except BaseException:
        print('The action "click" has been finished! ')
    finally:
        counter0 = counter0 - 1

# Get the hot movie ranking list
film_name = browser.find_elements_by_class_name('title')
film_name
print(type(film_name))
for name in film_name:
    print(name.text)

film_info = browser.find_elements_by_class_name('item')

myFile = xlwt.Workbook() # Create Excel workbook
workbook = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = workbook.add_sheet('test', cell_overwrite_ok=True)
# sheet.write( row, colume, 'content' ) related paraments
sheet.write(0, 0, '序号')  
sheet.write(0, 1, '电影名称')
sheet.write(0, 2, 'URL')
sheet.write(0, 3, '豆瓣评分')
sheet.write(0, 4, '海报链接')
sheet.write(0, 5, '导演')
sheet.write(0, 6, '编剧')
sheet.write(0, 7, '主演')
sheet.write(0, 8, '类型')
sheet.write(0, 9, '制片国家/地区')
sheet.write(0, 10, '上映日期')
sheet.write(0, 11, '片长')
sheet.write(0, 12, '其它名称')
sheet.write(0, 13, 'IMDB链接')
sheet.write(0, 14, '豆瓣标签')
sheet.write(0, 15, '剧情简介')
sheet.write(0, 16, '推荐')
workbook.save(r'C:\test1.xls')
print("成功创建Excel表格，暂存数据")
print("Successully create Excel workbook to store data.")

# 获取电影详细信息URL列表
# get the moive URL list
# 遍历所有电影
# traverse all the movies
filmURL = []
for element in film_info:    
    print(element.get_attribute('href'))
    filmURL.append(element.get_attribute('href'))
    
# the data structure to store the inof of a movie
myFilmDict = {
        'Number': 0,
        'filmName': '',
        'filmURL': '',
        'douMark': '',
        'posterURL': '',
        'director': '',
        'scriptwriter': '',
        'actor': '',
        'type': '',
        'area': '',
        'time': '',
        'runtime': '',
        'othername': '',
        'IMDB': '',
        'douTag': '',
        'summary': '',
        'recommendation': '',
        'recommendationURL': '',
        
        };

option = webdriver.ChromeOptions()  # Set the option
option.add_argument('headless')     # Headless browser
filmPage = webdriver.Chrome(chrome_options=option)           
        
#for index in range(len(filmURL)):
index = 0; # Set the beginning number of the movie
# You can change this number
# For example, you can get the infomation from the 10th movie.
while( index < (len(filmURL) - 1) ):
    try:
        # Use random generator to create an int number between 0 and 2 to delay 
        # In order to avoid the high frequency of sending requests to Douban Server
        # In order to avoid the 403/404 Forbidden on the Douban Firewall
        # If you delete this statement, the website probably will not be available for you.
        time.sleep(random.randint(0,2))
        
        timestamp1_ms = datetime.datetime.now()
        print(timestamp1_ms)
        
        # Input the URL on the browser
        filmPage.get(filmURL[index])
        index = index + 1
        
        myFilmDict['Number'] = index
        print("正在打印第" + str(index) + "部电影的信息" )
       
        # 电影名称 Film name
        name = filmPage.find_element_by_tag_name('h1')
        myFilmDict['filmName'] = name.text.strip()
        print(name.text.strip())
        myFilmDict['filmURL'] = filmPage.current_url
        # print(filmPage.current_url)
        
        # Douban Mark
        # 豆瓣评分
        mark = filmPage.find_element_by_class_name("rating_num")
        myFilmDict['douMark'] = mark.text.strip()
        # print( "豆瓣评分： " + mark.text.strip() )
        
        # Get the poster info
        # 获取海报信息
        posterURL = re.search('<img src="(.*?)" title="点击看更多海报"',filmPage.page_source,re.S)
        myFilmDict['posterURL'] = posterURL.group(1)
        # print("海报链接： " + posterURL.group(1))
       
        # Get the info of the directors, scriptwriters and actors
        # 提取导演、编剧、主演信息
        # 这里如果主演人数多，需要模拟鼠标点击事件，点击“更多”
        # 模拟人类点击事件，先定位按钮，再调用click()
        
        # This segment of exception catch code is used to click the button "More-actors."
        try:
            filmPage.find_element_by_class_name("more-actor").click()
        except BaseException:
            print('已经完成点击')
        
        # 开始匹配人员
        maker = filmPage.find_elements_by_class_name("attrs")
        myFilmDict['director'] = maker[0].text.strip()
        myFilmDict['scriptwriter'] = maker[1].text.strip()
        myFilmDict['actor'] = maker[2].text.strip()
        # print("导演： " + maker[0].text.strip())
        # print("编剧： " + maker[1].text.strip())
        # print("主演： " + maker[2].text.strip())
       
        # 提取电影基本信息
        film_basicDetail = re.search('<div id="info">(.*?)</div>',filmPage.page_source,re.S)
        # print(film_basicDetail.group())
        
        # film type
        #  匹配电影类型
        # <span class="pl">类型:</span> <span property="v:genre">剧情</span> / <span property="v:genre">犯罪</span><br />
        genre = re.findall('"v:genre">(.*?)</span>',film_basicDetail.group(), re.S)
        myFilmDict['type'] = "、".join(genre)
        # print("类型：", end="" )
        # for item in genre:
        #    print(item + '/', end="" )
        
        # Area
        # 匹配制片国家和地区
        # <span class="pl">制片国家/地区:</span> 美国<br />
        area = re.search('制片国家/地区:</span> (.*?)<br />',film_basicDetail.group(), re.S)
        myFilmDict['area'] = area.group(1)
        #print("\n制片国家/地区:"+area.group(1))
    
        # 上映时间 Reflection time
        # <span property="v:initialReleaseDate" content="1994-09-10(多伦多电影节)">1994-09-10(多伦多电影节)</span>
        releaseDate = re.findall('<span property="v:initialReleaseDate" content="(.*?)"',film_basicDetail.group(), re.S)  
        myFilmDict['time'] = "  ".join(releaseDate)
        # print("上映日期：", end="" )
        # for item in releaseDate:
        #     print(item + '/', end="" )
        
        # 片长 Runtime
        # <span class="pl">片长:</span> <span property="v:runtime" content="142">142分钟</span><br />
        runtime = re.search('"v:runtime" content="(.*?)"',film_basicDetail.group(), re.S)
        myFilmDict['runtime'] = runtime.group(1)
        # print("\n片长：" + runtime.group(1) + "分钟")
    
        # 获取其他片名 Othernames
        # <span class="pl">又名:</span> 月黑高飞(港) / 刺激1995(台) / 地狱诺言 / 铁窗岁月 / 消香克的救赎<br />
        try:
            otherName = re.search('又名:</span> (.*?)<br />',film_basicDetail.group(), re.S)
            myFilmDict['othername'] = otherName.group(1)
            # print("又名：" + otherName.group(1))
        except BaseException:
            print(str(BaseException)+"无法获取别名")
        
        # IMDB Link
        # <span class="pl">IMDb链接:</span> <a href="http://www.imdb.com/title/tt0111161" target="_blank" rel="nofollow">tt0111161</a><br />
        imdbLink = re.search('IMDb链接:</span> <a href="(.*?)" ',film_basicDetail.group(), re.S)
        myFilmDict['IMDB'] = imdbLink.group(1)
        # print("IMDB链接： " + imdbLink.group(1))
        
        # Douban tag for this movie
        # 豆瓣成员常用的标签
        # print("豆瓣成员常用的标签：", end="")        
        tag = filmPage.find_elements_by_class_name("tags-body")
       
        tagList = []    
        for item in tag:
            # print(item.text.strip())
            tagList.append(item.text.strip())
            
        myFilmDict['douTag'] = " ".join(tagList)
        
        # Summary of this movie
        # 获取剧情简介
        # print("\n剧情简介：",end="")
        try:
            if re.findall('展开全部',filmPage.page_source, re.S):
                #需要展开全部
                summary = filmPage.find_element_by_class_name("short").text
                myFilmDict['summary'] = summary
                # print(summary)
            else:
                #不需要展开全部
                summary = filmPage.find_element_by_id("link-report").text
                myFilmDict['summary'] = summary
                # print(summary)
        except BaseException:
            print(str(BaseException)+"无法获取剧情简介")
        
        # Other movies recommended by Douban
        # 喜欢这部电影的人也喜欢什么电影
        # print("喜欢这部电影的人也喜欢：")
        try:
            rec = filmPage.find_element_by_class_name("recommendations-bd")
            ddtag = rec.find_elements_by_tag_name("dd")
            #print(ddtag)
            nameList = []
            recURL_list = []
            for item in ddtag:        
                atag = item.find_element_by_tag_name("a") 
                rec_name = atag.text
                rec_url = atag.get_attribute('href')
                nameList.append(rec_name)
                recURL_list.append(rec_url)
                # print(rec_name)
                # print(rec_url)
            
            myFilmDict['recommendation'] = " ".join(nameList)
            myFilmDict['recommendationURL'] = " ".join(recURL_list)
        except NoSuchElementException:
            print("暂时没有数据")
        finally:
            print()
        

        timestamp2_ms = datetime.datetime.now()
        print(timestamp2_ms)
        
        # Append data in Excel 
        myExcel = xlrd.open_workbook(r'C:\test1.xls')
        myExcelnew = copy(myExcel)
        ws = myExcelnew.get_sheet(0)
        
        ws.write(index, 0, str(myFilmDict['Number']))  
        ws.write(index, 1, myFilmDict['filmName'])
        ws.write(index, 2, myFilmDict['filmURL'])
        ws.write(index, 3, myFilmDict['douMark'])
        ws.write(index, 4, myFilmDict['posterURL'])
        ws.write(index, 5, myFilmDict['director'])
        ws.write(index, 6, myFilmDict['scriptwriter'])
        ws.write(index, 7, myFilmDict['actor'])
        ws.write(index, 8, myFilmDict['type'])
        ws.write(index, 9, myFilmDict['area'])
        ws.write(index, 10, myFilmDict['time'])
        ws.write(index, 11, myFilmDict['runtime'])
        ws.write(index, 12, myFilmDict['othername'])
        ws.write(index, 13, myFilmDict['IMDB'])
        ws.write(index, 14, myFilmDict['douTag'])
        ws.write(index, 15, myFilmDict['summary'])
        ws.write(index, 16, myFilmDict['recommendation'])
        
        # print( str( sys.getsizeof(myFilmDict) ) + "字节等待写入磁盘" )
        
        myExcelnew.save(r'C:\test1.xls')
        
        timestamp3_ms = datetime.datetime.now()
        print(timestamp3_ms)
        print()
    
    except BaseException:
        print(BaseException)        

# 遍历所有电影完成
# Traversion finished!!        
        
browser.close() 
  
time2 = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) 

print(time1)
print(time2)


