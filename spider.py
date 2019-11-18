#!/usr/bin/python
#-*- coding: utf-8 -*-

import sys
import time
import urllib
import urllib.parse
import requests
import numpy as np
from bs4 import BeautifulSoup
from openpyxl import Workbook
from urllib import request
from urllib import error
import imp
imp.reload(sys)



#Some User Agents
hds=[{'User-Agent':'Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6'},\
{'User-Agent':'Mozilla/5.0 (Windows NT 6.2) AppleWebKit/535.11 (KHTML, like Gecko) Chrome/17.0.963.12 Safari/535.11'},\
{'User-Agent': 'Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.2; Trident/6.0)'}]

people=[]

def book_spider():
    page_num=1;
    builder_list=[]
    try_times=0
    
    while(page_num>0):
        url='https://www.mbawa.com/become-a-member/find-a-member/page/'+str(page_num)+'/?member_type=residential-builder&button&find-a-member=true'
        time.sleep(np.random.rand()*5) # Random waitling time to cover
        
        # Complex Version
        try:
            req = request.Request(url, headers=hds[page_num%len(hds)])
            source_code = request.urlopen(req).read()
            plain_text=str(source_code).replace('<br>',' ').replace('<br/>',' ').replace('<br />',' ')
        except (error.HTTPError, error.URLError) as e: 
            print (e)
            continue
        
        soup = BeautifulSoup(plain_text)
        list_soup = soup.find('ol', {'class': 'results table-list'})
        
        if list_soup==None: # or page_num>1:
            page_num=-1
            print('All result has been stored.')
            break

        for builder_info in list_soup.findAll('article'):
            name = builder_info.find('a', {'target':'_blank'}).string.strip()
            company_url = builder_info.find('a', {'target':'_blank'}).get('href')
            detail = get_detail(company_url, name)
            builder_list.append([name,detail[0],detail[1],detail[2],detail[3],detail[4],detail[5],detail[6],detail[7],detail[8]])
            try_times=0 # Set 0 when got valid information
        print ('Downloading Information From Page %d' % page_num)
        page_num+=1
    return builder_list


def get_detail(url, name):
    detail=[''] * 9
    global people

    try:
        req = request.Request(url, headers=hds[np.random.randint(0,len(hds))])
        source_code = request.urlopen(req).read()
        plain_text=str(source_code).replace('<br>','').replace('<br/>','').replace('<br />','').replace('\\n','')
    except (error.HTTPError, error.URLError) as e:
        print (e)

    soup = BeautifulSoup(plain_text)

    try:
        ul = soup.find('ul', {'class':'member-meta'}).find_all('li')
    except:
        print('Company data not found: '+name)
        return detail

    for li in ul:
        if(li.find('strong') == None):
            try:
                [s.extract() for s in li('strong')]
                content = li.string.strip()
            except:
                print('Wrong type address')
                continue
            else:
                detail[5]=content
        elif(li.find('strong').string == 'Builders Reg No: '):
            try:
                [s.extract() for s in li('strong')]
                content = li.string.strip()
            except:
                print('Wrong type Reg No.')
                continue
            else:
                detail[0]=content
        elif(li.find('strong').string == 'Areas:'):
            try:
                [s.extract() for s in li('strong')]
                content = li.string.strip()
            except:
                print('Wrong type area')
                continue
            else:
                detail[1]=content
        elif(li.find('strong').string == 'Contact:'):
            try:
                [s.extract() for s in li('strong')]
                content = li.string.strip()
            except:
                print('Wrong type contact')
                continue
            else:
                if content != '':
                    people.append([content,name])
                    detail[2]=content
        elif(li.find('strong').string == 'T:'):
            try:
                [s.extract() for s in li('strong')]
                content = li.string.strip()
            except:
                print('Wrong type T')
                continue
            else:
                detail[3]=content
        elif(li.find('strong').string == 'F:'):
            try:
                [s.extract() for s in li('strong')]
                content = li.string.strip()
            except:
                print('Wrong type F')
                continue
            else:
                detail[4]=content
    try:
        web = soup.find('a', {'class':'btn btn-primary btn-uppercase m-b-1'}).get('href')
        detail[6]=web
    except:
        print('No web')

    divl = soup.find_all('div', {'class':'m-b-3'})
    for item in divl:
        if(item.find('h3') != None and item.find('h3').string == 'Key Projects'):
            item_list=[]
            li=item.find_all('li')
            for i in li:
                try:
                    if i.string != None:
                        item_list.append(i.string)
                except:
                    continue
            p=item.find_all('p')
            for i in p:
                try:
                    if i.string != None:
                        item_list.append(i.string)
                except:
                    continue
            detail[7]=','.join(item_list)
        elif(item.find('h3') != None and item.find('h3').string == 'Awards'):
            item_list=[]
            li=item.find_all('li')
            for i in li:
                try:
                    if i.string != None:
                        item_list.append(i.string)
                except:
                    continue
            p=item.find_all('p')
            for i in p:
                try:
                    if i.string != None:
                        item_list.append(i.string)
                except:
                    continue
            detail[8]=','.join(item_list)

    stuff = soup.find('h4', {'class':'primary'})
    if stuff != None and len(stuff.find_next_siblings('ul')) != 0:
        for person in stuff.find_next_siblings('ul')[0].find_all('li'):
            try:
                if person.string.strip() != '':
                    people.append([person.string.strip(),name])
            except:
                continue
    elif stuff != None and stuff.find_next_siblings('p') != None:
        for person in stuff.find_next_siblings('p'):
            try:
                if person.string.strip() != '':
                    people.append([person.string.strip(),name])
            except:
                continue
        
    return detail

def print_builder_lists_excel(builder_lists):
    wb=Workbook(write_only=True)
    global people
    people = list(set([tuple(t) for t in people]))

    ws=wb.create_sheet('builder')
    ws2=wb.create_sheet('people')
    ws.append(['No.','Name','Reg No','Areas','Contact','T','F','Address','Web','Key Projects','Awards'])
    ws2.append(['No.','Name','Company'])
    count=1
    for bl in builder_lists:
       ws.append([count,bl[0],bl[1],bl[2],bl[3],bl[4],bl[5],bl[6],bl[7],bl[8],bl[9]])
       count+=1
    count=1
    for person in people:
       ws2.append([count,person[0],person[1]])
       count+=1
    save_path='builder list '+time.strftime("%Y-%m-%d", time.localtime())+'.xlsx'
    wb.save(save_path)




if __name__=='__main__':
    builder_lists=book_spider()
    print_builder_lists_excel(builder_lists)