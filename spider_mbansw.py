#!/usr/bin/python
#-*- coding: utf-8 -*-

import sys
import time
import urllib
import urllib.parse
import requests
import numpy as np
import imp
import re
from openpyxl import Workbook
from urllib import request
from urllib import error
from fake_useragent import UserAgent
imp.reload(sys)

# Fake User Agents
ua = UserAgent()

def builder_spider():
    page_num=0;
    builder_list=[['No.','Name','Company','Phone']]
    
    while(page_num>-1):
        url='https://www.mbansw.asn.au/find-a-master-builder?field_referral_categories_target_id=All&field_referral_areas_target_id=All&page='+str(page_num)
        time.sleep(np.random.rand()*5) # Random waitling time to cover
        
        # Try to get html code
        try:
            req = request.Request(url, headers={'User-Agent': ua.random})
            source_code = request.urlopen(req).read()
            plain_text=str(source_code).replace('<br>',' ').replace('<br/>',' ').replace('<br />',' ').replace('&amp;','&')
        except (error.HTTPError, error.URLError) as e: 
            print_builder_lists_excel(builder_list,page_num)
            page_num = -1
            print ('Access denied.')
            break

        company = re.findall(r'<h2 class="field-content">[^<]*</h2>',plain_text, re.S)
        name = re.findall(r'<div class="views-field views-field-field-user-last-name"><div class="field-content">[^<]*</div></div>',plain_text, re.S)
        phone = re.findall(r'<div class="views-field views-field-field-phone"><div class="field-content">[^<]*</div></div>',plain_text, re.S)
        
        if name == []:
            print_builder_lists_excel(builder_list,page_num)
            page_num = -1
            print ('No result left')
            break

        i=0
        while i < len(name):
            builder_list.append([len(builder_list),
                                 name[i].replace('<div class="views-field views-field-field-user-last-name"><div class="field-content">','').replace('</div></div>','').strip(),
                                 company[i].replace('<h2 class="field-content">','').replace('</h2>','').strip(),
                                 phone[i].replace('<div class="views-field views-field-field-phone"><div class="field-content">','').replace('</div></div>','').lstrip('Phone: ').strip()])
            i+=1
        print ('Downloading Information From Page %d' % page_num)
        page_num+=1
    return 1

def print_builder_lists_excel(builder_list,page_num):
    wb=Workbook(write_only=True)

    ws=wb.create_sheet('People')
    for bl in builder_list:
       ws.append([bl[0],bl[1],bl[2],bl[3]])
    save_path='BuilderFromMbansw '+time.strftime("%Y-%m-%d", time.localtime())+' page '+str(page_num)+'.xlsx'
    wb.save(save_path)

if __name__=='__main__':
    start = time.time()
    builder_lists=builder_spider()
    finish = time.time()
    time_elapsed = finish - start
    print('The code run {:.0f}m {:.0f}s'.format(time_elapsed // 60, time_elapsed % 60))