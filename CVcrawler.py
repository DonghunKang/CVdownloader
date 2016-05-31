# -*- coding: utf-8 -*-
"""
Created on Tue Jul 14 22:13:54 2015

@author: Administrator
"""

import urllib2
import requests as rs
import bs4
import xlsxwriter
import sys
import os 

# txt_to_list 메소드: 
# 논문제목 리스트 생성, papers_list.txt 파일 불러와서 리스트에 파일명 로딩
def txt_to_list(filename):
    lst=[]
    
    profs = open(filename,'r')
    for line in profs.readlines():
        #print type(line)
        lst.append(line)
    
    profs.close()
    
    lst = map(lambda s: s.strip(), lst)
    
    return lst
    
# find_cv_link 메소드:    
# Google search query(ex: https://www.google.co.kr/search?q=Jessica+Wachter+finance+cv+filetype:pdf) 
# 이용하여 다운로드할 CV의 pdf파일 주소를 리턴함
    
def find_cv_link(url):
    
    response = rs.get(url)
    html_content = response.text.encode(response.encoding)
    nav = bs4.BeautifulSoup(html_content)
    tmp = nav.findAll("h3", { "class" : "r" })[0]
    #tmp = <h3 class="r"><a href="/url?q=http://finance.wharton.upenn.edu/~jwachter/Wachtercv.pdf&amp;sa=U&amp;ved=0CBMQFjAAahUKEwiHvZmx4drGAhUXjo4KHWJVBiA&amp;usg=AFQjCNH9yJOwRs-5iIcUubTzldd5ImeujA" target="_blank"><b>CURRICULUM VITAE JESSICA</b> A. <b>WACHTER</b> April 2015 Address <b>...</b></a></h3> 
    pdf = str(tmp).split("&amp")[0].split("/url?q=")[1]
    #pdf = 'http://finance.wharton.upenn.edu/~jwachter/Wachtercv.pdf'
    return pdf 
    


# download_file 메소드:
# 1. 저자명, 2. 다운로드할 cv의 url 입력
def download_file(author, download_url):
    indicator=0
    try:
        response = urllib2.urlopen(download_url, timeout = 2)
        file = open(author+".pdf", 'wb')
        file.write(response.read())
        file.close()
        print(author+" Completed")
        indicator=1
        return indicator
        
        
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]      
        print(exc_type, fname, exc_tb.tb_lineno)
        return indicator
'''
    except urllib2.URLError, e:
        raise 
    except urllib2.HTTPError, err:
       if err.code == 404:
           print "!!! 404 not found !!!"
       else:
           raise            
'''           

    
'''
try:
   urllib2.urlopen("some url")
except urllib2.HTTPError, err:
   if err.code == 404:
       <whatever>
   else:
       raise    

'''



## prof_list, url_list 생성

prof_list = txt_to_list('profs_list.txt')   
#prof_list = txt_to_list('profs_list_tmp.txt')  
#prof_list = txt_to_list('profs_list.txt')
query_list =[]
url_list = []
download_list =[]
header = 'https://www.google.com/search?q='
footer = '+cv+finance+filetype:pdf'

for prof in prof_list:
    prof_name='+'.join(prof.strip().lower().split(' '))
    #prof_name = 'a+variance+decomposition+for+stock+returns'
    query = header+prof_name+footer
    print query
    query_list.append(query)
    #query="https://www.google.co.kr/search?q=Jessica+Wachter+finance+cv+filetype:pdf"
    url = find_cv_link(query)
    print url
    url_list.append(str(url))
    #download pdf
    download_list.append(download_file(prof, url))




### 최종 결과를 xlsx파일에 기록 ###

workbook = xlsxwriter.Workbook('result.xlsx')
worksheet = workbook.add_worksheet()

row = 0
col = 0

for i in range(len(prof_list)):
    worksheet.write(row,col,prof_list[i])
    worksheet.write(row,col+1,query_list[i])
    worksheet.write(row,col+2,url_list[i])
    worksheet.write(row,col+3,download_list[i])
    row = row + 1

workbook.close()










    
    