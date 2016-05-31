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

def lstToSoup(lst):
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
    

def tagSplit(tag,strSplit):
    return strToSoup(str(tag).split(strSplit))    

# linktoReport 메소드만 있음
# input: KIND 보고서 링크
# output: 
    
def linkToReport(link):
    #print link
    nav_tmp=code_to_html('',link,'')
    nav=bs(str(nav_tmp).split('SECTION-1')[-1])
    title_tags= nav.find_all(class_="TABLE")
    
    tmp =[]
    pages =[]    #pages 리스트 참고사항: 1:발행회사 2:발행주식 3:보고개요 4: 세부거래사항 5: 최대주주 소유현황
    header=[]
    trans_Content=[]
    status_Content=[]        
    
    
    
    for title_tag in title_tags:
        page = [str(title_tag)]
        elem = next_element(title_tag)
        #print elem
        while elem and elem.name != 'table': #while 문에 있는 태그 꼭 수정해야 함!
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
    firmInfo = tagsToList(pages[0].find_all(class_="TD"))[1::4]
    eventDate=re.search('200\d.+[01]\d.+\d{2}',link).group(0).replace('/','-')
    stockIssued=tagsToList(tagSplit(pages[1],'<tr>')[1].find_all(class_="TD"))
    
    header = firmInfo+[eventDate]+stockIssued
    
    ## 4. 세부거래사항의 내용 표로 만들기 위한 전처리
    ## 세부거래사장의 헤더(변동사항, 성명, 최대주주 및 발행회사와의 관계, 국적)와 내용 합쳐서 
    ## trans_Content 리스트에 이중 리스트 형태로 넣었음.
    ## [[보유자1, 변경사항 1-1 ~ 1-4], [보유자2, 변경사항 2-1 ~ 2-3], ... ]형태
    # 200x년 xx월 xx일 형식으로 저장
    #pages[2].find_all(text=re.compile("200\d.{2}[01]\d.{2}\d{2}"))[-1].encode("utf-8")
    
    ## 5. 최대주주등 주식소유현황 내용 표로 만들기 위한 전처리
    ## 주식소유현황 헤더(변동사항, 성명, 최대주주 및 발행회사와의 관계, 국적)와 내용 합쳐서 
    ## status_Content의 데이터 형태: 이중 리스트
    trans_Content = []
    status_Content = []
    
    if len(pages) == 2:
        table= tagSplit(pages[1],'<tr>')[2:-1]
    else:
        table = tagSplit(pages[2],'<tr>')[1:-1]
    
    for row in table:
        if len(row.find_all(class_='TD')): # 한 행에 TD 클래스가 1개 이상 있을 경우
            #print row.find(class_='TD').text
            if row.find(class_='TD').text.count('200'):
                tmpList=tagsToList(row.find_all(class_='TD'))         
                p = re.compile("(20\d{2}).+([01]\d).+(\d{2})")
                d=p.match(tmpList[0])
                date=d.group(1)+'-'+d.group(2)+'-'+d.group(3)
                tmpList[0]=date
                #print date
                trans_Content.append(tmpList)
            if row.find_all(class_='TD')[-1].text != '-': 
                ## status_Content 리스트 작성하는 부분:
                # 한 행의 마지막 TD 클래스가 - 값이 아닐 경우, 즉 값이 있을 경우를 status에 추가
                tmpList2=tagsToList(row.find_all(class_='TD'))
                del tmpList2[5:-6]
                del tmpList2[:2]
                #print tmpList2
                status_Content.append(tmpList2)
    
    report = (header, trans_Content, status_Content)

    return report



#%% txt파일 웹에서 로딩
import time
start_time = time.time()

def makeReports():
    reports = []
    crashed_links=[]
    link_report = txt_to_list('links_42392_54982.txt')
    num_reports = len(link_report)
    idx0 = 0            
    for link in link_report:
        print link        
        try:        
            reports.append(linkToReport(link))
            idx0 = idx0 + 1
            print 'appending link: ', idx0 ,'of',num_reports,'done'
        except:
            print 'CRASHED!!! Check URL OF: ', idx0 ,'of',num_reports,'done'
            idx0 = idx0 + 1
            crashed_links.append(link)
            pass
        continue
        

    f = open("crash_report_42392_54982.txt", 'w') 
    for links in crashed_links: 
        data = links+'\n'
        f.write(data) 
    f.close()

    return reports

reports = makeReports()



# 엑셀파일로 기록

code = txt_to_list('codes_42392_54982.txt')    
link = txt_to_list('links_42392_54982.txt')
workbook = xlsxwriter.Workbook('test_output_42392_54982.xlsx')
worksheet = workbook.add_worksheet()

row = 1
col = 1

## header --> report[0]
## trans_Content --> report[1]
## status_Content --> report[2]
idx = 0

crashed_links=[]

for report in reports:
    
    try:
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
    except:
        print 'CRASHED!!! Check URL'
        crashed_links.append(link[idx])
        idx = idx + 1
        

print("--- running time : %s seconds ---" % (time.time() - start_time))
workbook.close()

f = open("crash_report_two_tables_42392_54982.txt", 'w') 
for links in crashed_links: 
    data = links+'\n'
    f.write(data) 
f.close()


#%% test !!!!
#link='http://kind.krx.co.kr/external/2001/01/05/000056/20010105000098/2001010500009806.htm'
#link='http://kind.krx.co.kr/external/2001/03/12/000004/20010312000006/2001031200000606.htm'
#link='http://kind.krx.co.kr/external/2000/10/05/000044/20001005000071/2000100500007106.htm'
link='http://kind.krx.co.kr/external/2000/11/27/000023/20001127000037/2000112700003706.htm'


nav_tmp=code_to_html('',link,'')
nav=bs(str(nav_tmp).split('SECTION-1')[-1])
title_tags= nav.find_all(class_="TABLE")

tmp =[]
pages =[]    #pages 리스트 참고사항: 1:발행회사 2:발행주식 3:보고개요 4: 세부거래사항 5: 최대주주 소유현황
header=[]
trans_Content=[]
status_Content=[]        
crashed_links = []


for title_tag in title_tags:
    page = [str(title_tag)]
    elem = next_element(title_tag)
    #print elem
    while elem and elem.name != 'table': #while 문에 있는 태그 꼭 수정해야 함!
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
#firmName=pages[0].find_all(class_="TD")[0].text.encode("utf-8")
#firmCode=pages[0].find_all(class_="TD")[1].text.encode("utf-8")
#commStockIssued=pages[1].find_all(class_="TD")[0].text.encode("utf-8")


    

firmInfo = tagsToList(pages[0].find_all(class_="TD"))[1::4]
eventDate=re.search('200\d/[01]\d/\d{2}',link).group(0).replace('/','-')

stockIssued=tagsToList(tagSplit(pages[1],'<tr>')[1].find_all(class_="TD"))

header = firmInfo+[eventDate]+stockIssued

#pages[2].find_all(class_="TD",colspan=1,rowspan=3)[-1].text




#print header

 #
## 4. 세부거래사항의 내용 표로 만들기 위한 전처리
## 세부거래사장의 헤더(변동사항, 성명, 최대주주 및 발행회사와의 관계, 국적)와 내용 합쳐서 
## trans_Content 리스트에 이중 리스트 형태로 넣었음.
## [[보유자1, 변경사항 1-1 ~ 1-4], [보유자2, 변경사항 2-1 ~ 2-3], ... ]형태
# 200x년 xx월 xx일 형식으로 저장
#pages[2].find_all(text=re.compile("200\d.{2}[01]\d.{2}\d{2}"))[-1].encode("utf-8")

## 5. 최대주주등 주식소유현황 내용 표로 만들기 위한 전처리
## 주식소유현황 헤더(변동사항, 성명, 최대주주 및 발행회사와의 관계, 국적)와 내용 합쳐서 
## status_Content의 데이터 형태: 이중 리스트
trans_Content = []
status_Content = []

if len(pages) == 2:
    table= tagSplit(pages[1],'<tr>')[2:-1]
else:
    table = tagSplit(pages[2],'<tr>')[1:-1]

for row in table:
    if len(row.find_all(class_='TD')): # 한 행에 TD 클래스가 1개 이상 있을 경우
        #print row.find(class_='TD').text
        if row.find(class_='TD').text.count('200'):
            tmpList=tagsToList(row.find_all(class_='TD'))         
            p = re.compile("(20\d{2}).+([01]\d).+(\d{2})")
            d=p.match(tmpList[0])
            date=d.group(1)+'-'+d.group(2)+'-'+d.group(3)
            tmpList[0]=date
            #print date
            trans_Content.append(tmpList)
        if row.find_all(class_='TD')[-1].text != '-': 
            ## status_Content 리스트 작성하는 부분:
            # 한 행의 마지막 TD 클래스가 - 값이 아닐 경우, 즉 값이 있을 경우를 status에 추가
            tmpList2=tagsToList(row.find_all(class_='TD'))
            del tmpList2[5:-6]
            del tmpList2[:2]
            #print tmpList2
            status_Content.append(tmpList2)
            
#%%
# test -- Excel 쓰기
workbook = xlsxwriter.Workbook('test_output_t.xlsx')
worksheet = workbook.add_worksheet()

row = 1
col = 1
idx = 0
## header --> report[0]
## trans_Content --> report[1]
## status_Content --> report[2]
try:
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
except:
    print 'CRASHED!!! Check URL'
    crashed_links.append(link)
    pass
        
workbook.close()












