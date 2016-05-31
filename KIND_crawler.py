# -*- coding: utf-8 -*-
"""
Created on Tue Jul 14 22:13:54 2015

@author: KDH
"""
#%%

import requests as rs
import bs4
import xlsxwriter

# txt_to_list 메소드: 
# 논문제목 리스트 생성, papers_list.txt 파일 불러와서 리스트에 파일명 로딩
def txt_to_list(filename):
    lst=[]
    
    profs = open(filename,'r')
    for line in profs.readlines():
        #print type(line)
        lst.append(line)
    profs.close()
    
    lst = map(lambda s: s.strip(), lst) #/n 제거
    
    return lst


# code_to_html 메소드:
# header+code+footer 내용에 대해 request하고 해당 url의 내용을 nav 객체로 return함

def code_to_html(header, code, footer):
    url=header+code+footer
    response = rs.get(url)
    html_content = response.text.encode('utf-8')
    nav = bs4.BeautifulSoup(html_content, from_encoding="utf-8")
    return nav

# extract_url 메소드:
# acptno --> docNo --> url_ext(최종경로) 추출하는 메소드

def extract_url(acptno):
    ## 1. 읽어온 acptno로 docNo 추출
    header = 'http://kind.krx.co.kr/common/disclsviewer.do?method=search&acptno='
    footer = ''
    nav_doc=code_to_html(header,acptno,footer)
    # nav_doc.find_all('option',selected="selected")[0]['value']  ### option value
    docNo=nav_doc.find_all('option',selected="selected")[0]['value'].split("|")[0]
    
    ## 2. docNo에 해당하는 주소로 들어가서 최종 경로 추출
    header2 = 'http://kind.krx.co.kr/common/disclsviewer.do?method=searchContents&docNo='
    nav_ext=code_to_html(header2,docNo,footer)
    url_ext=str(nav_ext.find('script')).split(',')[1].strip("'")
    #url_ext='http://kind.krx.co.kr/external/2015/07/07/000222/20150629000613/99602.htm'
    return url_ext



#### STEPS ####
# 1. txt파일에서 acptno읽어오기
# 2. 읽어온 acptno로 docNo 추출
# 3. docNo에 해당하는 주소로 들어가서 최종 경로 추출
# 4. 최종경로 (보고서 파일이 있는 웹 주소)에서 필요한 데이터 추출



# main 문
def main():
    #acptno_list=txt_to_list('acptno_tmp.txt')
    ##acptno_list=txt_to_list('acptno100.txt')
    #num=1
    #for i in range(4):
    #idx=i+1
    #filename='acpt_code_'+str(idx)+'.txt'
    #acptno_list=txt_to_list(filename)
    acptno_list=txt_to_list('acptno3.txt')
    #urlext_list=[]
    #print urlext_list
    
    ## 엑셀 파일로 기록
    xlsname='ext_result3.xlsx'
    workbook = xlsxwriter.Workbook(xlsname)
    worksheet = workbook.add_worksheet()
    
    #worksheet.write(0,0,"num")
    worksheet.write(0,0,"acptno")
    worksheet.write(0,1,"URL")
    
    row = 1
    col = 0
    
    for j in range(len(acptno_list)):
        #worksheet.write(row,col,num)            
        worksheet.write(row,col,acptno_list[j])
        url = extract_url(acptno_list[j])
        worksheet.write(row,col+1,url)
        row = row + 1
        
        if row % 50 == 0:
            print 'row :', row, url
        
        
    workbook.close()


main()

#%% 불러온 urlext_list의 html 내용을 파일 형태로 정리


a=urlext_list[0]
nav=code_to_html('',a,'')
l=nav.find_all('span')
for tag in l:
    print tag.text



#%%
f = open("tmp.txt", 'w')

for title in titles: 
    data = str(title)+'\n' 
    f.write(data) 

f.close()


'''
## 1. 읽어온 acptno로 docNo 추출
acptno="20150721000075"
header = 'http://kind.krx.co.kr/common/disclsviewer.do?method=search&acptno='
footer = ''
nav_doc=code_to_html(header,acptno,footer)
# nav_doc.find_all('option',selected="selected")[0]['value']  ### option value
docNo=nav_doc.find_all('option',selected="selected")[0]['value'].split("|")[0]

## 2. docNo에 해당하는 주소로 들어가서 최종 경로 추출
header2 = 'http://kind.krx.co.kr/common/disclsviewer.do?method=searchContents&docNo='
nav_ext=code_to_html(header2,docNo,footer)
url_ext=str(nav_ext.find('script')).split(',')[1].strip("'")
#url_ext='http://kind.krx.co.kr/external/2015/07/07/000222/20150629000613/99602.htm'



r = rs.get(url_ext)
html_content = r.text.encode(response.encoding)
nav_final = bs4.BeautifulSoup(r.content)
'''





"""

url=header+content+footer
response = rs.get(url_ext)
html_content = response.text.encode(response.encoding)
nav = bs4.BeautifulSoup(html_content)
tmp = nav.findAll("h3", { "class" : "r" })[0]
#tmp = <h3 class="r"><a href="/url?q=http://finance.wharton.upenn.edu/~jwachter/Wachtercv.pdf&amp;sa=U&amp;ved=0CBMQFjAAahUKEwiHvZmx4drGAhUXjo4KHWJVBiA&amp;usg=AFQjCNH9yJOwRs-5iIcUubTzldd5ImeujA" target="_blank"><b>CURRICULUM VITAE JESSICA</b> A. <b>WACHTER</b> April 2015 Address <b>...</b></a></h3> 
pdf = str(tmp).split("&amp")[0].split("/url?q=")[1]
#pdf = 'http://finance.wharton.upenn.edu/~jwachter/Wachtercv.pdf'
return pdf 














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
        
"""        
        
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

    
    
'''
a
nav
code_to_html
a
nav=code_to_html('',a,'')
nav
nav=code_to_html('',a,'')
nav
a=urlext_list[0]
nav=code_to_html('',a,'')
nav
s=bs4.BeautifulSoup(nav)
s.find_all('table')
bs4
BeautifulSoup(nav)
nav
nd_all('table')
le')
nav.find_all('table')
nav.find_all('span')
l=nav.find_all('span')
l
'''