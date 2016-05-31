# -*- coding: utf-8 -*-
"""
Created on Tue Jul 14 22:13:54 2015

@author: KDH
"""

import re
import requests as rs
import bs4
import xlsxwriter
# sys.setdefaultencoding() does not exist, here!
import sys
reload(sys)  # Reload does the trick!
sys.setdefaultencoding('UTF8')

def code_to_html(header, code, footer):
    url=header+code+footer
    response = rs.get(url)
    #html_content = response.text.encode('utf-8')
    #nav = bs4.BeautifulSoup(html_content, from_encoding="utf-8")
    html_content = response.text.encode(response.encoding)
    nav = bs4.BeautifulSoup(html_content)
    return nav
   

def next_element(elem):
    while elem is not None:
        # Find next element, skip NavigableString objects
        elem = elem.next_sibling
        if hasattr(elem, 'name'):
            return elem

def printVal(html):
    tmpVal= html.find_all(class_="xforms_input")
    for tmp in tmpVal:
        print tmp.text

def strToSoup(lst):
    idx = 0    
    for item in lst:
        lst[idx]=bs4.BeautifulSoup(item)
        idx = idx + 1
    return lst
        
def tagsToList(tags):
    lst = []    
    for tag in tags:
        lst.append(tag.text)
    return lst

def chunks_with_header(l,h,n):
    n = max(1, n)
    return [h+l[i:i + n] for i in range(0, len(l), n)]
    
def txt_to_list(filename):
    lst=[]
    
    profs = open(filename,'r')
    for line in profs.readlines():
        #print type(line)
        lst.append(line)
    profs.close()
    
    lst = map(lambda s: s.strip(), lst) #/n 제거
    
    return lst



# linktoReport 메소드만 있음
    
def linkToReport(link):
    
    nav=code_to_html('',link,'')
    title_tags= nav.find_all(style=re.compile("11pt"))
    
    
    pages =[]    #pages 리스트 참고사항: 1:발행회사 2:발행주식 3:보고개요 4: 세부거래사항 5: 최대주주 소유현황
    header=[]
    trans_Content=[]
    status_Content=[]        
    
    for title_tag in title_tags:
        page = [str(title_tag)]
        elem = next_element(title_tag)
        #print elem
        while elem and elem.name != 'span':
            page.append(str(elem))
            elem = next_element(elem)
            #print len(page)
        #print '!!!'
        pages.append('\n'.join(page))
    
    tmp=0
    for page in pages:    
        page_html =  bs4.BeautifulSoup(page)    
        pages[tmp] = page_html
        tmp = tmp+1
        
    
    ## 첫번째 헤더 추출하기 (회사명, 종목코드, 공시일자, 발행보통주식수)
    firmName=pages[0].find_all(class_="xforms_input")[0].text.encode("utf-8")
    firmCode=pages[0].find_all(class_="xforms_input")[1].text.encode("utf-8")
    eventDate=pages[2].find_all(class_="xforms_input", style=re.compile("center"))[-1].text.encode("utf-8")
    commStockIssued=tagsToList(pages[1].find_all(class_="xforms_input"))[0:3]
    #commStockIssued=tagsToList(pages[1].find_all(class_="xforms_input"))[0]
    header = [firmName,firmCode,eventDate]+commStockIssued    
    
    #eventDate=pages[2].find_all(class_="xforms_input", style=re.compile("center"))[-1].text.encode("utf-8")
    #commStockIssued=pages[1].find_all(class_="xforms_input")[0].text.encode("utf-8")
    #header = [firmName,firmCode,eventDate,commStockIssued]
    
    #print header
    
     
    ## 4. 세부거래사항의 내용 표로 만들기 위한 전처리
    ## 세부거래사장의 헤더(변동사항, 성명, 최대주주 및 발행회사와의 관계, 국적)와 내용 합쳐서 
    ## trans_Content 리스트에 이중 리스트 형태로 넣었음.
    ## (보유자1, 변경사항 1-1 ~ 1-4), (보유자2, 변경사항 2-1 ~ 2-3), ... 형태
    
    
    tables_content=str(pages[3]).split("<tbody>")
    tables_content=tables_content[1:]
    
    #re.compile("[12][90]\d\d.[01]\d.\d{2}").split(tables_content[0])[0]
    
    for table in tables_content:   
        header_len=re.compile("[12][90]\d\d.[01]\d.\d{2}").split(tables_content[0])[0].count('<tr>')
        #print 'HEADER LENGTH :',header_len
        #re.split(r'[12][90]\d\d.[01]\d.\d{2}',table)[0]
        rows=table.split("<tr>")
        h=tagsToList(bs4.BeautifulSoup('\n'.join(rows[:header_len])).find_all(class_="xforms_input"))
        if header_len < 5:
            h = h + ['-']*5
        content=tagsToList(bs4.BeautifulSoup('\n'.join(rows[header_len:])).find_all(class_="xforms_input"))
        countRows=rows[header_len].count("xforms")        
        #print len(content)
        rows_content=chunks_with_header(content,h,countRows)
        trans_Content=trans_Content+rows_content #한 row당 7개
    
        
    ## 5. 최대주주등 주식소유현황 내용 표로 만들기 위한 전처리
    ## 주식소유현황 헤더(변동사항, 성명, 최대주주 및 발행회사와의 관계, 국적)와 내용 합쳐서 
    ## status_Content 리스트에 이중 리스트 형태로 넣었음.
        
    table_status=str(pages[4]).split("<tbody>")[1]
    rows_status=table_status.split("<tr>")
    content_status=tagsToList(bs4.BeautifulSoup('\n'.join(rows_status[3:-1])).find_all(class_="xforms_input"))
    countRows2=rows_status[-2].count("xforms")
    
    status_Content=chunks_with_header(content_status,[],countRows2) #한 row당 17개
    
    report = (header, trans_Content, status_Content)

    return report

# 7, 17 바꿔야 함


#%% txt파일 웹에서 로딩
import time
start_time = time.time()

def makeReports():
    reports = []
    crashed_links=[]
    link_report = txt_to_list('links_1_21940.txt')
    num_reports = len(link_report)
    idx0 = 0     
       
    for link in link_report:
        print link        
        try:        
            reports.append(linkToReport(link))
            idx0 = idx0 + 1
            print 'appending link: ', idx0 ,'of',num_reports,'done'
        except:
            print e
            print 'CRASHED!!! Check URL OF: ', idx0 ,'of',num_reports,'done'
            idx0 = idx0 + 1
            crashed_links.append(link)
        
    

    f = open("crash_report_1_21940.txt", 'w') 
    for links in crashed_links: 
        data = links+'\n'
        f.write(data) 
    f.close()            
            
            
    return reports
    
reports = makeReports()
    
# 엑셀파일로 기록
code = txt_to_list('codes_1_21940.txt')    
workbook = xlsxwriter.Workbook('test_output_1_21940.xlsx')
worksheet = workbook.add_worksheet()

row = 1
col = 1

## header --> report[0]
## trans_Content --> report[1]
## status_Content --> report[2]
idx = 0



for report in reports:
    
    
    for i in range(len(report[1])):
                    
        for k in range(len(report[0])):
            worksheet.write(row+i,col+k,report[0][k])
        
        worksheet.write(row+i,0,code[idx])    
        worksheet.write(row+i,col+len(report[0]),'transaction_'+str(len(report[1][0])))  
    
        for j in range(len(report[1][0])):
            worksheet.write(row+i,col+len(report[0])+1+j,report[1][i][j])
    
    row = row + len(report[1])
    
    for i in range(len(report[2])):
        for k in range(len(report[0])):
            worksheet.write(row+i,col+k,report[0][k])
        worksheet.write(row+i,0,code[idx])
        worksheet.write(row+i,col+len(report[0]),'stock_status_'+str(len(report[2][0])))    
    
        for j in range(len(report[2][0])):
            worksheet.write(row+i,col+len(report[0])+1+j,report[2][i][j])
    
    row = row + len(report[2])
    idx = idx + 1

    

print("--- running time : %s seconds ---" % (time.time() - start_time))
workbook.close()





#%% test !!!!
link = 'http://kind.krx.co.kr/external/2009/04/29/000013/20090428016898/99602.htm'
#link = 'http://kind.krx.co.kr/external/2012/07/20/000409/20120718000622/99602.htm' #외국인 혼자 주주
#link = 'http://kind.krx.co.kr/external/2012/07/20/000384/20120720000053/99602.htm' #주식소유현황 갯수 17에서 변수로 바꾸기 
nav=code_to_html('',link,'')
title_tags= nav.find_all(style=re.compile("11pt"))


pages =[]    #pages 리스트 참고사항: 1:발행회사 2:발행주식 3:보고개요 4: 세부거래사항 5: 최대주주 소유현황
header=[]
trans_Content=[]
status_Content=[]        

for title_tag in title_tags:
    page = [str(title_tag)]
    elem = next_element(title_tag)
    #print elem
    while elem and elem.name != 'span':
        page.append(str(elem))
        elem = next_element(elem)
        #print len(page)
    #print '!!!'
    pages.append('\n'.join(page))

tmp=0
for page in pages:    
    page_html =  bs4.BeautifulSoup(page)    
    pages[tmp] = page_html
    tmp = tmp+1
    

## 첫번째 헤더 추출하기 (회사명, 종목코드, 공시일자, 발행보통주식수)
firmName=pages[0].find_all(class_="xforms_input")[0].text.encode("utf-8")
firmCode=pages[0].find_all(class_="xforms_input")[1].text.encode("utf-8")
eventDate=pages[2].find_all(class_="xforms_input", style=re.compile("center"))[-1].text.encode("utf-8")
commStockIssued=tagsToList(pages[1].find_all(class_="xforms_input"))[0:3]
#commStockIssued=tagsToList(pages[1].find_all(class_="xforms_input"))[0]
header = [firmName,firmCode,eventDate]+commStockIssued

print header
#%%
 
## 4. 세부거래사항의 내용 표로 만들기 위한 전처리
## 세부거래사장의 헤더(변동사항, 성명, 최대주주 및 발행회사와의 관계, 국적)와 내용 합쳐서 
## trans_Content 리스트에 이중 리스트 형태로 넣었음.
## (보유자1, 변경사항 1-1 ~ 1-4), (보유자2, 변경사항 2-1 ~ 2-3), ... 형태


tables_content=str(pages[3]).split("<tbody>")
tables_content=tables_content[1:]

#re.compile("[12][90]\d\d.[01]\d.\d{2}").split(tables_content[0])[0]

for table in tables_content:   
    header_len=re.compile("[12][90]\d\d.[01]\d.\d{2}").split(tables_content[0])[0].count('<tr>')
    print 'HEADER LENGTH :',header_len
    #re.split(r'[12][90]\d\d.[01]\d.\d{2}',table)[0]
    rows=table.split("<tr>")
    h=tagsToList(bs4.BeautifulSoup('\n'.join(rows[:header_len])).find_all(class_="xforms_input"))
    if header_len < 5:
        h = h + ['-']*5
    content=tagsToList(bs4.BeautifulSoup('\n'.join(rows[header_len:])).find_all(class_="xforms_input"))
    countRows=rows[header_len].count("xforms")
    print 'countRows: ',countRows
    #print len(content)
    rows_content=chunks_with_header(content,h,countRows)
    trans_Content=trans_Content+rows_content #한 row당 7개

    
## 5. 최대주주등 주식소유현황 내용 표로 만들기 위한 전처리
## 주식소유현황 헤더(변동사항, 성명, 최대주주 및 발행회사와의 관계, 국적)와 내용 합쳐서 
## status_Content 리스트에 이중 리스트 형태로 넣었음.
    
table_status=str(pages[4]).split("<tbody>")[1]
rows_status=table_status.split("<tr>")
content_status=tagsToList(bs4.BeautifulSoup('\n'.join(rows_status[3:-1])).find_all(class_="xforms_input"))
countRows2=rows_status[-2].count("xforms")
print 'rownum:',countRows2
status_Content=chunks_with_header(content_status,[],countRows2) #한 row당 17개

    
workbook = xlsxwriter.Workbook('test_output.xlsx')
worksheet = workbook.add_worksheet()

row = 1
col = 1
idx = 0
## header --> report[0]
## trans_Content --> report[1]
## status_Content --> report[2]
   
    

#row - col 순서로 추가하는 이중 for문
for i in range(len(trans_Content)):
    for k in range(len(header)):
        worksheet.write(row+i,col+k,header[k])
    worksheet.write(row+i,col+len(header),'transaction')    
    for j in range(len(trans_Content[0])):
        worksheet.write(row+i,col+len(header)+1+j,trans_Content[i][j])

row=row+len(trans_Content)

for i in range(len(status_Content)):
    for k in range(len(header)):
        worksheet.write(row+i,col+k,header[k])
    worksheet.write(row+i,col+len(header),'stock_status')    
    for j in range(len(status_Content[0])):
        worksheet.write(row+i,col+len(header)+1+j,status_Content[i][j])
    
workbook.close()














#%%            
    
workbook = xlsxwriter.Workbook('test_output.xlsx')
worksheet = workbook.add_worksheet()

row = 1
col = 1

## header --> report[0]
## trans_Content --> report[1]
## status_Content --> report[2]
idx = 0

    
    

    #row - col 순서로 추가하는 이중 for문
    for i in range(len(trans_Content)):
        for k in range(len(header)):
            worksheet.write(row+i,col+k,header[k])
        
        worksheet.write(row+i,col+len(header),'transaction')    
    
        for j in range(len(trans_Content[0])):
            worksheet.write(row+i,col+len(header)+1+j,trans_Content[i][j])
    
    row=row+len(trans_Content)
    
    for i in range(len(status_Content)):
        for k in range(len(header)):
            worksheet.write(row+i,col+k,header[k])
        
        worksheet.write(row+i,col+len(header),'stock_status')    
    
        for j in range(len(status_Content[0])):
            worksheet.write(row+i,col+len(header)+1+j,status_Content[i][j])
        
    workbook.close()


main()


#%%

link='http://kind.krx.co.kr/external/2015/07/16/000204/20150715000130/99602.htm'
nav=code_to_html('',link,'')
title_tags= nav.find_all(style=re.compile("11pt"))


pages =[]    #pages 리스트 참고사항: 1:발행회사 2:발행주식 3:보고개요 4: 세부거래사항 5: 최대주주 소유현황
header=[]
trans_Content=[]
status_Content=[]        

for title_tag in title_tags:
    page = [str(title_tag)]
    elem = next_element(title_tag)
    #print elem
    while elem and elem.name != 'span':
        page.append(str(elem))
        elem = next_element(elem)
        #print len(page)
    #print '!!!'
    pages.append('\n'.join(page))

tmp=0
for page in pages:    
    page_html =  bs4.BeautifulSoup(page)    
    pages[tmp] = page_html
    tmp = tmp+1
    

## 첫번째 헤더 추출하기 (회사명, 종목코드, 공시일자, 발행보통주식수)
firmName=pages[0].find_all(class_="xforms_input")[0].text.encode("utf-8")
firmCode=pages[0].find_all(class_="xforms_input")[1].text.encode("utf-8")
eventDate=pages[2].find_all(class_="xforms_input", style=re.compile("center"))[-1].text.encode("utf-8")
commStockIssued=pages[1].find_all(class_="xforms_input")[0].text.encode("utf-8")

header = [firmName,firmCode,eventDate,commStockIssued]
#print header

 
## 4. 세부거래사항의 내용 표로 만들기 위한 전처리
## 세부거래사장의 헤더(변동사항, 성명, 최대주주 및 발행회사와의 관계, 국적)와 내용 합쳐서 
## trans_Content 리스트에 이중 리스트 형태로 넣었음.
## (보유자1, 변경사항 1-1 ~ 1-4), (보유자2, 변경사항 2-1 ~ 2-3), ... 형태

tables_content=str(pages[3]).split("<tbody>")
tables_content=tables_content[1:]
for table in tables_content:   
    rows=table.split("<tr>")
    h=tagsToList(bs4.BeautifulSoup('\n'.join(rows[:5])).find_all(class_="xforms_input"))
    content=tagsToList(bs4.BeautifulSoup('\n'.join(rows[5:])).find_all(class_="xforms_input"))
    print len(content)
    trans_Content=trans_Content+chunks_with_header(content,h,7) #한 row당 7개

    
## 5. 최대주주등 주식소유현황 내용 표로 만들기 위한 전처리
## 주식소유현황 헤더(변동사항, 성명, 최대주주 및 발행회사와의 관계, 국적)와 내용 합쳐서 
## status_Content 리스트에 이중 리스트 형태로 넣었음.
    
table_status=str(pages[4]).split("<tbody>")[1]
rows_status=table_status.split("<tr>")
content_status=tagsToList(bs4.BeautifulSoup('\n'.join(rows_status[3:-1])).find_all(class_="xforms_input"))
status_Content=chunks_with_header(content_status,[],17) #한 row당 17개


report = (header, trans_Content, status_Content)









#%%


for i in range(len(trans_Content)):
    print "tansaction",i+1

for j in range(len(status_Content)):
    print "status",j+1



workbook = xlsxwriter.Workbook('test_output.xlsx')
worksheet = workbook.add_worksheet()

row = 1
col = 0

cal = [['January','February','March'],['April','May','June']]
for i in range(len(cal)):
    for j in range(len(cal[0])):
        worksheet.write(row+i,col+j,cal[i][j])
    


workbook.close()



'''
for i in range(len(acptno_list)):
    worksheet.write(row,col,acptno_list[i])
    worksheet.write(row,col+1,urlext_list[i])
    row = row + 1
'''
workbook.close()



#%%

def pagesToTables(pages):
    tables=str(pages[3]).split("<tbody>")
    tables=tables[1:] #1번째 element 삭제
    idx = 0
    for table in tables:
        tables[idx]= bs4.BeautifulSoup(table)
        idx = idx + 1
    return tables

tables = pagesToTables(pages)
print len(tables)


f = open("tmp.txt", 'w')

for title in titles: 
    data = str(title)+'\n' 
    f.write(data) 

f.close()


tmpVal= pages[0].find_all(class_="xforms_input")
for tmp in tmpVal:
    print tmp.text

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

#acptno_list=txt_to_list('acptno_tmp.txt')
##acptno_list=txt_to_list('acptno100.txt')
acptno_list=txt_to_list('acptno_300.txt')
urlext_list=[]
'''
for acptno in acptno_list:
    url_ext=extract_url(acptno)
    print url_ext
    urlext_list.append(url_ext)
'''
#print urlext_list

## 엑셀 파일로 기록

workbook = xlsxwriter.Workbook('ext_result.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write(0,0,"acptno")
worksheet.write(0,1,"URL")

row = 1
col = 0

for i in range(len(acptno_list)):
    worksheet.write(row,col,acptno_list[i])
    url = extract_url(acptno_list[i])
    print row, url
    worksheet.write(row,col+1,url)
    row = row + 1
'''
for i in range(len(acptno_list)):
    worksheet.write(row,col,acptno_list[i])
    worksheet.write(row,col+1,urlext_list[i])
    row = row + 1
'''
workbook.close()




    