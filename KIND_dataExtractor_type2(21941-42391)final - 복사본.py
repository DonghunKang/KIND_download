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

def bs(text):
    return bs4.BeautifulSoup(text)

# linktoReport 메소드만 있음
    
def linkToReport(link):
    #print link
    nav=code_to_html('',link,'')
    title_tags= nav.find_all(class_="P")
    
    
    pages =[]    #pages 리스트 참고사항: 1:발행회사 2:발행주식 3:보고개요 4: 세부거래사항 5: 최대주주 소유현황
    header=[]
    trans_Content=[]
    status_Content=[]        
    
    
    for title_tag in title_tags:
        page = [str(title_tag)]
        elem = next_element(title_tag)
        #print elem
        while elem and elem.name != 'p': #while 문에 있는 태그 꼭 수정해야 함!
            page.append(str(elem))
            elem = next_element(elem)
            #print len(page)
        #print '!!!'
        pages.append('\n'.join(page))
    
    tmp=0
    for page in pages:    
        #print tmp    
        page_html =  bs4.BeautifulSoup(page)    
        pages[tmp] = page_html
        tmp = tmp+1
        
    
    ## 첫번째 헤더 추출하기 (회사명, 종목코드, 공시일자, 발행보통주식수)
    firmName=pages[0].find_all(class_="TD")[0].text.encode("utf-8")
    #firmCode=pages[0].find_all(class_="TD")[1].text.encode("utf-8")
    #eventDate=pages[2].find_all(text=re.compile("200\d.{2}[01]\d.{2}\d{2}"))[-1].encode("utf-8")
    if pages[2].find_all(text=re.compile("200\d.{2}[01]\d.{2}\d{2}")) != []:
        eventDate=pages[2].find_all(text=re.compile("200\d.{2}[01]\d.{2}\d{2}"))[-1].encode("utf-8")
    else:
        if pages[3].find_all(text=re.compile("200\d.{2}[01]\d.{2}\d{2}"))[0] != []:
            eventDate=pages[3].find_all(text=re.compile("200\d.{2}[01]\d.{2}\d{2}"))[0]
        else:
            eventDate='-'
    
    #commStockIssued=pages[1].find_all(class_="TD")[0].text.encode("utf-8")
    #header = [firmName,'-',eventDate,commStockIssued]
    commStockIssued=tagsToList(pages[1].find_all(class_="TD"))[0:3]
    header = [firmName,'-',eventDate]+commStockIssued
    #pages[2].find_all(class_="TD",colspan=1,rowspan=3)[-1].text
    
    
    #print header
    
     
    ## 4. 세부거래사항의 내용 표로 만들기 위한 전처리
    ## 세부거래사장의 헤더(변동사항, 성명, 최대주주 및 발행회사와의 관계, 국적)와 내용 합쳐서 
    ## trans_Content 리스트에 이중 리스트 형태로 넣었음.
    ## [[보유자1, 변경사항 1-1 ~ 1-4], [보유자2, 변경사항 2-1 ~ 2-3], ... ]형태
    # 200x년 xx월 xx일 형식으로 저장
    #pages[2].find_all(text=re.compile("200\d.{2}[01]\d.{2}\d{2}"))[-1].encode("utf-8")
    
    header_len=0
    tables_content=str(pages[3]).split("<table")
    tables_content=tables_content[1:]
    
    for table in tables_content:
        rows=table.split("<tr>")[1:]
        for i in range(len(rows)):
            if rows[i].count("TD") == 0:
                header_len=i
                break
        #print 'HEADER LENGTH :',header_len
        h=tagsToList(bs4.BeautifulSoup('\n'.join(rows[:header_len])).find_all(class_="TD"))
        content=tagsToList(bs4.BeautifulSoup('\n'.join(rows[header_len:])).find_all(class_="TD"))
        countRows=rows[-1].count("TD")
        #print 'countRows: ',countRows
        #print len(content)
        rows_content=chunks_with_header(content,h,countRows)
        trans_Content=trans_Content+rows_content #한 row당 7개
    
    ## 5. 최대주주등 주식소유현황 내용 표로 만들기 위한 전처리
    ## 주식소유현황 헤더(변동사항, 성명, 최대주주 및 발행회사와의 관계, 국적)와 내용 합쳐서 
    ## status_Content의 데이터 형태: 이중 리스트
    
    status_Content =[]
    rows_status=strToSoup(str(pages[4]).split('<tr>'))
    for row in rows_status:
        if str(row).count('<td') == str(row).count('TD'):
            status_Content=status_Content+[tagsToList(row.find_all(class_='TD'))]
    
    status_Content=status_Content[1:]
    
    report = (header, trans_Content, status_Content)

    return report

# 7, 17 바꿔야 함


#%% txt파일 웹에서 로딩
import time
start_time = time.time()

def makeReports():
    reports = []
    crashed_links=[]
    link_report = txt_to_list('links_21941_42391.txt')
    num_reports = len(link_report)
    idx0 = 0            
    for link in link_report:
        try:        
            reports.append(linkToReport(link))
            idx0 = idx0 + 1
            print 'appending link: ', idx0 ,'of',num_reports,'done'
        except IndexError:
            print 'CRASHED!!! Check URL OF: ', idx0 ,'of',num_reports,'done'
            idx0 = idx0 + 1
            crashed_links.append(link)
            pass
        continue
        

    f = open("crash_report_21941_42391.txt", 'w') 
    for links in crashed_links: 
        data = links+'\n'
        f.write(data) 
    f.close()

    return reports

reports = makeReports()


#

# 엑셀파일로 기록
code = txt_to_list('codes_21941_42391.txt')    
workbook = xlsxwriter.Workbook('test_output_21941_42391.xlsx')
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
#link = 'http://kind.krx.co.kr/external/2008/08/20/000151/20080820000289/2008082000028904.htm'
#link = 'http://kind.krx.co.kr/external/2006/12/20/000127/20061220000204/2006122000020404.htm'
#link='http://kind.krx.co.kr/external/2005/02/24/000280/20050224000501/2005022400050106.htm'
link='http://kind.krx.co.kr/external/2005/11/22/000048/20051122000080/2005112200008006.htm'

nav=code_to_html('',link,'')
title_tags= nav.find_all(class_="P")


pages =[]    #pages 리스트 참고사항: 1:발행회사 2:발행주식 3:보고개요 4: 세부거래사항 5: 최대주주 소유현황
header=[]
trans_Content=[]
status_Content=[]        


for title_tag in title_tags:
    page = [str(title_tag)]
    elem = next_element(title_tag)
    #print elem
    while elem and elem.name != 'p': #while 문에 있는 태그 꼭 수정해야 함!
        page.append(str(elem))
        elem = next_element(elem)
        #print len(page)
    #print '!!!'
    pages.append('\n'.join(page))

tmp=0
for page in pages:    
    #print tmp    
    page_html =  bs4.BeautifulSoup(page)    
    pages[tmp] = page_html
    tmp = tmp+1
   

## 첫번째 헤더 추출하기 (회사명, 종목코드, 공시일자, 발행보통주식수)
firmName=pages[0].find_all(class_="TD")[0].text.encode("utf-8")
#firmCode=pages[0].find_all(class_="TD")[1].text.encode("utf-8")
if pages[2].find_all(text=re.compile("200\d.{2}[01]\d.{2}\d{2}")) != []:
    eventDate=pages[2].find_all(text=re.compile("200\d.{2}[01]\d.{2}\d{2}"))[-1].encode("utf-8")
else:
    if pages[3].find_all(text=re.compile("200\d.{2}[01]\d.{2}\d{2}"))[0] != []:
        eventDate=pages[3].find_all(text=re.compile("200\d.{2}[01]\d.{2}\d{2}"))[0]
    else:
        eventDate='-'
    
#commStockIssued=pages[1].find_all(class_="TD")[0].text.encode("utf-8")
commStockIssued=tagsToList(pages[1].find_all(class_="TD"))[0:3]
header = [firmName,'-',eventDate]+commStockIssued

#pages[2].find_all(class_="TD",colspan=1,rowspan=3)[-1].text


#print header

#%%  
## 4. 세부거래사항의 내용 표로 만들기 위한 전처리
## 세부거래사장의 헤더(변동사항, 성명, 최대주주 및 발행회사와의 관계, 국적)와 내용 합쳐서 
## trans_Content 리스트에 이중 리스트 형태로 넣었음.
## [[보유자1, 변경사항 1-1 ~ 1-4], [보유자2, 변경사항 2-1 ~ 2-3], ... ]형태
# 200x년 xx월 xx일 형식으로 저장
#pages[2].find_all(text=re.compile("200\d.{2}[01]\d.{2}\d{2}"))[-1].encode("utf-8")

header_len=0
tables_content=str(pages[3]).split("<table")
tables_content=tables_content[1:]

for table in tables_content:
    rows=table.split("<tr>")[1:]
    for i in range(len(rows)):
        if rows[i].count("TD") == 0:
            header_len=i
            break
    #print 'HEADER LENGTH :',header_len
    h=tagsToList(bs4.BeautifulSoup('\n'.join(rows[:header_len])).find_all(class_="TD"))
    content=tagsToList(bs4.BeautifulSoup('\n'.join(rows[header_len:])).find_all(class_="TD"))
    countRows=rows[-1].count("TD")
    #print 'countRows: ',countRows
    #print len(content)
    rows_content=chunks_with_header(content,h,countRows)
    trans_Content=trans_Content+rows_content #한 row당 7개

## 5. 최대주주등 주식소유현황 내용 표로 만들기 위한 전처리
## 주식소유현황 헤더(변동사항, 성명, 최대주주 및 발행회사와의 관계, 국적)와 내용 합쳐서 
## status_Content의 데이터 형태: 이중 리스트

status_Content =[]
rows_status=strToSoup(str(pages[4]).split('<tr>'))
for row in rows_status:
    if str(row).count('<td') == str(row).count('TD'):
        status_Content=status_Content+[tagsToList(row.find_all(class_='TD'))]

status_Content=status_Content[1:]

# test -- Excel 쓰기
workbook = xlsxwriter.Workbook('test_output_t.xlsx')
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












