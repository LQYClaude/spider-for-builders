#!/usr/bin/python
#-*- coding: utf-8 -*-

import sys
import time
import urllib
import urllib.parse
import requests
import numpy as np
import imp
from bs4 import BeautifulSoup
from openpyxl import Workbook
from urllib import request
from urllib import error
from fake_useragent import UserAgent
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
imp.reload(sys)

# Fake User Agents
ua = UserAgent()
hd = ua.random

def builder_spider():
    page_num=1
    builder_list=[['No.','Name','Address','Suburb','State','Postcode','Logo','YellowPage Detail','Rating','Reviews','Phone','Web','Email','latitude','longitude','Show Case','Awards']]
    
    while(page_num>0):
        url='https://www.yellowpages.com.au/search/listings?clue=Builders+%26+Building+Contractors&pageNumber='+str(page_num)+'&referredBy=www.yellowpages.com.au&&eventType=pagination'
        time.sleep(np.random.rand()*10) # Random waitling time to cover
        
        # Grab page by chrome driver
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument(f'user-agent={hd}')
        browser = webdriver.Chrome(chrome_options=chrome_options)

        try:
            browser.get(url)
        except:
            print_builder_lists_excel(builder_list,page_num-1)
            pring('Access denied.')
            break

        soup = BeautifulSoup(browser.page_source, 'html.parser')
        try:
            list_result = soup.body.find('div', {'class': 'search-results search-results-data listing-group'})
        except:
            print_builder_lists_excel(builder_list,page_num-1)
            pring('Access denied.')
            break
        else:
            if list_result == None:
                print_builder_lists_excel(builder_list,page_num-1)
                print('Access denied.')
                break

        if list_result.find('div', {'class': 'listing listing-search listing-data', 'data-is-top-of-list': 'true'}) == None or list_result.find_all('div', {'class': 'listing listing-search listing-data', 'data-is-top-of-list': 'false'}) == None:
            print_builder_lists_excel(builder_list,page_num-1)
            print('All result has been stored.')
            page_num=-1
            break

        [s.extract() for s in list_result('span')]
        for builder_info in list_result.find_all('div', {'class': 'listing listing-search listing-data', 'data-is-top-of-list': 'false'}):
            name = builder_info.get('data-full-name')                                           # 1.company name
            address = builder_info.get('data-full-address')                                     # 2.address
            suburb = builder_info.get('data-suburb')                                            # 3.suburb
            state = builder_info.get('data-state')                                              # 4.state
            postcode = builder_info.get('data-postcode')                                        # 5.postcode

            try:
                logo = builder_info.find('img', {'class':'listing-logo enhanced-logo'}).get('src').lstrip('//')  # 6.logo url
            except:
                logo = ''
                print('No logo')

            try:
                yellow_url = builder_info.find('a', {'class':'image logo'}).get('href')         # 7.yellow page detail url
            except:
                yellow_url = ''
                print('No yellow page link')

            rating = builder_info.get('data-omniture-average-rating')                           # 8.rating
            reviews = builder_info.get('data-total-reviews')                                    # 9.reviews

            try:
                phone = builder_info.find('a', {'title':'Phone'}).get('href').lstrip('tel:')    # 10.phone
            except:
                phone = ''
                print('No phone')

            try:
                web = builder_info.find('a', {'class':'contact contact-main contact-url'}).get('href')                     # 11.web
            except:
                web = ''
                print('No web')

            try:
                email = builder_info.find('a', {'class':'contact contact-main contact-email'}).get('data-email')           # 12.email
            except:
                email = ''
                print('No email')

            try:
                latitude = builder_info.find('p', {'class':'listing-address mappable-address'}).get('data-geo-latitude')   # 13.latitude
                longitude = builder_info.find('p', {'class':'listing-address mappable-address'}).get('data-geo-longitude') # 14.longitude
            except:
                latitude = ''
                longitude = ''
                print('No latitude and longitude')

            # 15.show case url
            try:
                show_url = builder_info.find('a', {'class':'button transparent-background blue-text grey-88-border small-text-size promo-tile-link'}).get('href')
            except:
                show_url = ''
                print('No show case')

            # 16.awards
            awards_list = []
            for ul in builder_info.find_all('a', {'class': 'usp-awards-link-to-bpp'}):
                for li in ul.find_all('li'):
                    try:
                        awards_list.append(li.string.strip())
                    except:
                        continue
            awards=','.join(awards_list)

            builder_list.append([len(builder_list),name,address,suburb,state,postcode,logo,yellow_url,rating,reviews,phone,web,email,latitude,longitude,show_url,awards])
        print ('Downloading Information From Page %d' % page_num)
        page_num+=1
    print_builder_lists_excel(builder_list,page_num-1)
    return 1

def print_builder_lists_excel(builder_list,page_num):
    wb=Workbook(write_only=True)
    ws=wb.create_sheet('builder')
    for bl in builder_list:
       ws.append([bl[0],bl[1],bl[2],bl[3],bl[4],bl[5],bl[6],bl[7],bl[8],bl[9],bl[10],bl[11],bl[12],bl[13],bl[14],bl[15],bl[16]])
    save_path='builder list '+time.strftime("%Y-%m-%d", time.localtime())+' page '+str(page_num)+'.xlsx'
    wb.save(save_path)

if __name__=='__main__':
    builder_lists=builder_spider()