# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
import os
import sys
from selenium import webdriver
import xlwt
import time

driver = webdriver.Chrome()
vocabulary= ['区域卫生信息']
def url_define(i,kw):
	url = 'http://search.ccgp.gov.cn/bxsearch?searchtype=2&page_index=' + str(i) + '&bidSort=0&buyerName=&projectId=&pinMu=0&bidType=7&dbselect=bidx&kw=' + kw + '&start_time=2017%3A05%3A15&end_time=2018%3A05%3A15&timeType=6&displayZone=&zoneId=&pppStatus=0&agentName='
	return url

def href(url):
	driver.get(url)
	String = driver.page_source
	soup_0 = BeautifulSoup(String, 'lxml')
	if soup_0.find(name = 'title').string == '安全验证':
		time.sleep(20)
		driver.get(url)
		String = driver.page_source
		soup_0 = BeautifulSoup(String, 'lxml')	
	page_number = int(soup_0.find(name = 'p', attrs = {'style': 'float:left'}).contents[3].string)//20 + 1
	address_list = []
	for obj in soup_0.find(name = 'ul', attrs = {'class': 'vT-srch-result-list-bid'}).descendants:
		if (obj.name == 'a' and 'target' in list(obj.attrs.keys())):
			address_list.append(obj.attrs['href'])
	return (address_list, page_number)

def fetch(url):
	driver.get(url)
	String = driver.page_source
	soup_0 = BeautifulSoup(String, 'lxml')
	if soup_0.find(name = 'title').string == '安全验证':
		time.sleep(20)
		driver.get(url)
		String = driver.page_source
		soup_0 = BeautifulSoup(String, 'lxml')	
	bid_name = soup_0.find(name = 'title').string
	bid_date = soup_0.find(name = 'div', attrs = {'class': 'vF_detail_main'}).find(text = '中标日期').parent.next_sibling.string
	bid_unit = soup_0.find(name = 'div', attrs = {'class': 'vF_detail_main'}).find(text = '采购单位').parent.next_sibling.string
	bid_location = soup_0.find(name = 'div', attrs = {'class': 'vF_detail_main'}).find(text = '行政区域').parent.next_sibling.string
	bid_pro = soup_0.find(name = 'div', attrs = {'class': 'vF_detail_main'}).find(text = '采购项目名称').parent.next_sibling.string
	find_sth = soup_0.find(name = 'div', attrs = {'class': 'vF_detail_main'}).find(text = '总中标金额')
	if find_sth is None:
		return (bid_name, '详见公告正文')
	else:
		return (bid_name, find_sth.parent.next_sibling.string, bid_date, bid_unit, bid_location, bid_pro)

w = xlwt.Workbook()
sheet = w.add_sheet('Sheet 1')
sheet.write(0, 0, '中标公告名称')
sheet.write(0, 1, '总中标金额')
sheet.write(0, 2, '中标日期')
sheet.write(0, 3, '采购单位')
sheet.write(0, 4, '行政区域')
sheet.write(0, 5, '采购项目名称')
row = 0
for word in vocabulary:
	for i in range(1, href(url_define(1, word))[1] + 1):
		url = url_define(i, word)
		for j in href(url)[0]:
			row += 1
			sheet.write(row, 0, fetch(j)[0])
			sheet.write(row, 1, fetch(j)[1])
			sheet.write(row, 2, fetch(j)[2])
			sheet.write(row, 3, fetch(j)[3])
			sheet.write(row, 4, fetch(j)[4])
			sheet.write(row, 5, fetch(j)[5])
w.save('区域卫生信息化.xls')