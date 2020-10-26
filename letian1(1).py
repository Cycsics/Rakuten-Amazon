# -*- coding: utf-8 -*-
# Author ：CYCSICS
import time
import tkinter
import os
from tkinter import *
from tkinter.filedialog import askdirectory
from tkinter import messagebox
from tkinter import ttk
from tkinter import scrolledtext
from tkinter import END
from bs4 import BeautifulSoup
import datetime
import threading
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager


options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors')
# options.add_argument("--user-data-dir="+r"C:\\Users\\Administrator\\AppData\\Local\\Google\\Chrome\\User Data")
dr = webdriver.Chrome(chrome_options=options)
#ChromeDriverManager().install() 在windows环境下自动安装chromedriver
# '/usr/lib/chromium-browser/chromedriver'   树莓派环境下chromium位置
# dr = webdriver.Chrome('/usr/lib/chromium-browser/chromedriver')
# dr = webdriver.Chrome(ChromeDriverManager().install())
# dr.implicitly_wait(30)
handles = dr.window_handles
root = Tk()
root.geometry('150x450')
root.title("'楽天'处理器")
listbox_input = tkinter.scrolledtext.ScrolledText(root,width = 100)
listbox_output = tkinter.scrolledtext.ScrolledText(root,width = 200)
message_url = ''
path = StringVar()
lt_info_url = "https://order-rp.rms.rakuten.co.jp/order-rb/individual-order-detail-sc/init?orderNumber="
lt_list_url = 'https://order-rp.rms.rakuten.co.jp/order-rb/order-list-sc/init?&SEARCH_MODE=1&ORDER_PROGRESS=100'
lt_list_check_url = "https://order-rp.rms.rakuten.co.jp/order-rb/order-list-sc/"
amazon_url = "https://www.amazon.co.jp/gp/css/order-history?ref_=nav_orders_first"
amazon_check_url = 'https://www.amazon.co.jp'
records_names = []
lt_account1 = "bubu2016aA1P"
pwd1 = "benetech0927"
lt_account2 = "957320820@qq.com"
pwd2 = "RoseMerry818"
amazon_account = "sato@bjgbf.com"
amazon_pwd = "91sf7ac2"
wait_time = 2
credit_card_number = "4649882195642418"
order_infomation_id = ''
address_3line=[]

##################################################################################################
#
#句柄以及窗口切换流程：
#1、获取所有句柄（handles = dr.window_handles）
#2、循环切换对应句柄的窗口（dr.switch_to.window(handle)）
#3、获取当前窗口的url（dr.current_url）
#4、判断当前url是哪个网址并储存进对应的变量（考虑直接每次匹配对应url进行切换，不保存变量）
#5、切换对应变量储存的窗口，例（dr.switch_to.window(lt_list_handle)）
#
##################################################################################################


#自动化操作
def autoctl():
	lt_login()
	Amazon_login()
	
#操控浏览器
def lt_login():
	# set little time stop and big time stop for viewing changes
	dr.get("https://glogin.rms.rakuten.co.jp/?sp_id=1")
	WebDriverWait(dr,20).until(EC.element_to_be_clickable((By.ID,"rlogin-username-ja")))
	#第一个登陆流程
	dr.find_element_by_id("rlogin-username-ja").clear()
	dr.find_element_by_id("rlogin-password-ja").clear()
	dr.find_element_by_id("rlogin-username-ja").clear()
	dr.find_element_by_id("rlogin-username-ja").send_keys(lt_account1)
	dr.find_element_by_id("rlogin-password-ja").send_keys(pwd1)
	time.sleep(3)
	dr.find_element_by_xpath("//button[@name = 'submit']").click()
	# above = dr.find_element_by_xpath("/html/body/div[2]/main/div/section[1]/form/p[4]/button")
	# ActionChains(dr).move_to_element(above).click()
	#  
	WebDriverWait(dr,15).until(EC.element_to_be_clickable((By.NAME,"submit")))
#def second_login():
	#第二个登陆流程
	dr.find_element_by_id("rlogin-username-2-ja").clear()
	dr.find_element_by_id("rlogin-password-2-ja").clear()
	dr.find_element_by_id("rlogin-username-2-ja").clear()
	dr.find_element_by_id("rlogin-username-2-ja").send_keys(lt_account2)
	dr.find_element_by_id("rlogin-password-2-ja").send_keys(pwd2)
	dr.find_element_by_xpath("//button[@name = 'submit']").click()
	#  
	WebDriverWait(dr,15).until(EC.presence_of_element_located((By.NAME,"submit")))
	dr.find_element_by_xpath("//button[@name = 'submit']").click()
	#  
	WebDriverWait(dr,15).until(EC.presence_of_element_located((By.XPATH,"//*[@id='confirm']/p/button")))
	dr.find_element_by_xpath("//button[@type = 'submit']").click()
	# WebDriverWait(dr,15).until(EC.presence_of_element_located((By.XPATH,"//*[@id='confirm']/p/button")))
	# dr.find_element_by_xpath("//button[@type = 'submit']").click()
	# WebDriverWait(dr,15).until(EC.presence_of_element_located((By.XPATH,"//*[@id='confirm']/p/button")))
	# dr.find_element_by_xpath("//button[@type = 'submit']").click()
	# 
	# WebDriverWait(dr,15).until(EC.presence_of_element_located((By.XPATH,"//*[@id='mm_gg002_more002']")))
	# 点击进入订单界面
	dr.get("https://order-rp.rms.rakuten.co.jp/order-rb/order-list-sc/init?&SEARCH_MODE=1&ORDER_PROGRESS=100")
	# 点击进入待发送信件订单界面
	above = dr.find_element_by_xpath('//*[@id="filter1"]/ul/li[4]/a/h5/span')
	dr.execute_script("arguments[0].click();", above)
	# dr.find_element_by_xpath("//*[@id='filter1']/ul/li[4]/a").click()
	# dr.find_element_by_xpath("//*[@id='filter1']/ul/li[4]/a").send_keys(Keys.ENTER)
	time.sleep(3)

#管理发送邮件
def send_message():
	try:
		#句柄判断
		amazon_handle = ''
		lt_list_handle = ''
		lt_info_handle = ''
		handles = dr.window_handles
		for handle in handles:
			dr.switch_to.window(handle)
			curl = dr.current_url
			#判断当前url是否为亚马逊url
			if curl.find(amazon_check_url) != -1:
				amazon_handle = handle
			#判断当前url是否为乐天订单列表url
			elif curl.find(lt_list_check_url) != -1:
				lt_list_handle = handle
			#判断当前url是否为乐天订单详情url
			elif curl.find(lt_info_url) != -1:
				lt_info_handle = handle
		#切换到乐天订单列表详情
		dr.switch_to.window(lt_list_handle)
		#利用重新登录网址达成刷新目的(直接刷新窗口会有刷不出来新单的情况)
		dr.get(lt_list_url)
		#切换到待发送新建订单的窗口
		print("点击待发送")
		WebDriverWait(dr,10,0.5,ignored_exceptions=None).until(
			EC.element_to_be_clickable((By.XPATH,'//*[@id="filter1"]/ul/li[3]/a/h5/span'))
		)
		time.sleep(5)
		above = dr.find_element_by_xpath('//*[@id="filter1"]/ul/li[4]/a/h5/span')
		dr.execute_script("arguments[0].click();", above)
		# above = "$('div[id=filter1] ul li[3] h5 span').click()"
		# dr.execute_script(above)
		# above = dr.find_element_by_xpath('//*[@id="filter1"]/ul/li[3]/a/h5/span')
		# dr.find_element_by_xpath("//*[@id='filter1']/ul/li[4]/a").click()
		# ActionChains(dr).click(above).perform()
		#解码网页
		source = dr.page_source.encode('GBK', 'ignore')
		strsource = source.decode('GBK','strict')
		soup = BeautifulSoup(strsource,"html.parser")
		#定位有页码的列表
		ul = soup.find('ul',class_ = 'pagination')
		listbox_output.insert(END,'开始发信\n')
		#执行发送欢迎信件的函数
		print("开始发送欢迎信件")
		send_welcome_message()
		#执行发送发货信件的函数
		print("开始发送发货信件")
		send_deliver_message()
		#发信完毕切换到订单列表的窗口
		dr.switch_to.window(lt_list_handle)
		#利用重新登录网址达成刷新目的(直接刷新窗口会有刷不出来新单的情况)
		dr.get(lt_list_url)
		#点击到待发送信件的订单窗口
		WebDriverWait(dr,10,0.5,ignored_exceptions=None).until(
			EC.element_to_be_clickable((By.XPATH,'//*[@id="filter1"]/ul/li[4]/a/h5/span'))
		)
		time.sleep(2)
		# above = "$('div[id=filter1] ul li[4] h5 span').click()"
		# dr.execute_script(above)
		above = dr.find_element_by_xpath('//*[@id="filter1"]/ul/li[4]/a/h5/span')
		dr.execute_script("arguments[0].click();", above)
		# dr.find_element_by_xpath("//*[@id='filter1']/ul/li[4]/a").click()
		# dr.find_element_by_xpath("//*[@id='filter1']/ul/li[4]/a").send_keys(Keys.ENTER)
		# dr.find_element_by_xpath('//*[@id="filter1"]/ul/li[4]/a/h5/span').click()
		# dr.find_element_by_xpath("//li[@data-order-list-filter-tab='発送待ち']").click()
		Select(dr.find_element_by_xpath('//*[@id="ddlDisplay"]')).select_by_value("500")
		# ActionChains(dr).click(above).perform()
	except BaseException as e:
		print(e)
		if str(e).find("no such window: window was already closed"):
			ha = dr.window_handles
			dr.switch_to.window(lt_list_handle)	


#发货信件
#发货信件操作流程
##################################################################################################
#
#1、没有发送过发货信件
#判断流程：
#找到属性为"('span',attrs={"data-original-title": '発送のご案内'}))"的字段并判断class中是否存在'disabled'字样
#已发送发货信的class："rms-icon icon light rms-icon-square-mail-2-mini "
#未发送发货信的class："rms-icon icon light rms-icon-square-mail-2-mini disabled"
#
##################################################################################################
#
#2、在订单列表界面有"未入力"字样
#判断流程：
#获取class为"rms-text-alert rms-text-bold"的标签并获取文本内容，判断文本内容是否为"未入力"
#
##################################################################################################
#
#整体流程：
#1、订单列表查找符合条件1（没有发送过发货信件）的订单id
#2、判断id是否满足条件2（在订单列表界面有"未入力"字样），满足则加入list
#2、查找完毕以后点击进入订单（当前应获取所有句柄，并循环切换句柄对应窗口判断当前url来判断句柄对应的窗口是哪个）
#3、获取句柄之后切换到（lt_info_handle）对应的窗口
#4、不符合条件2时切换回（lt_list_handle）对应的窗口并重新进行循环
#5、符合条件2进入发信url进行发信
#
##################################################################################################
###			发货邮件获取条件需要重写，句柄操作有bug		###
def send_deliver_message():
	#重写
	amazon_handle = ''
	lt_list_handle = ''
	lt_info_handle = ''
	handles = dr.window_handles
	for handle in handles:
		dr.switch_to.window(handle)
		curl = dr.current_url
		#判断当前url是否为乐天订单列表url
		if curl.find(lt_list_check_url) != -1:
			lt_list_handle = handle
	dr.switch_to.window(lt_list_handle)
	WebDriverWait(dr,10,0.5,ignored_exceptions=None).until(
			EC.element_to_be_clickable((By.XPATH,'//*[@id="filter1"]/ul/li[4]/a/h5/span'))
		)
	time.sleep(5)
	# above = "$('div[id=filter1] ul li[4] h5 span').click()"
	# dr.execute_script(above)
	# above = dr.find_element_by_xpath('//*[@id="filter1"]/ul/li[4]/a/h5/span')
	# dr.find_element_by_xpath('//*[@id="filter1"]/ul/li[4]/a/h5/span').click()
	# ActionChains(dr).click(above).perform()
	above = dr.find_element_by_xpath('//*[@id="filter1"]/ul/li[4]/a/h5/span')
	dr.execute_script("arguments[0].click();", above)
	# dr.find_element_by_xpath("//*[@id='filter1']/ul/li[4]/a").click()
	# dr.find_element_by_xpath("//*[@id='filter1']/ul/li[4]/a").send_keys(Keys.ENTER)
	# dr.find_element_by_xpath("//li[@data-order-list-filter-tab='発送待ち']").click()
	time.sleep(5)
	source = dr.page_source.encode('GBK', 'ignore')
	strsource = source.decode('GBK','strict')
	soup = BeautifulSoup(strsource,"html.parser")
	tbodys = soup.find_all('tbody',class_ = 'rms-content-order-list-item rms-list-item rms-item-order-inprogress')
	tot = 0
	ids = []
	print("开始检查符合发信条件的订单")
	# print(tbodys)
	for tbody in tbodys:
		span1 = str(tbody.find('span',attrs={"data-original-title": '発送のご案内'}))
		span2 = str(tbody.find('span',attrs={'class','rms-text-alert rms-text-bold'}))
		print('span:',span2)
		loc = -1
		try:
			loc = span2.find('未入力')
			# print('未入力：Loc = ',loc)
		except:
			loc = -1
			pass
			
		if span1.find('disabled') != -1:
			print("disabled")
			if loc == -1:
				print('ID append')
				ids.append(tbody.find('a',target = 'order_detail').get_text())
	t = 1
	if len(ids) == 0:
		return
	for i in reversed(ids):
		print("当前正在发送发货邮件，订单号为%s\n"%i)
		# listbox_output.insert(END,'订单号为  %s  发货信件已发送\n'%i)
		url1 = 'https://order-rp.rms.rakuten.co.jp/order-rb/individual-order-detail-sc/init?orderNumber='+i
		if t == 1:
			dr.switch_to.window(lt_list_handle)
			time.sleep(2)
			above = dr.find_element_by_link_text(i)
			dr.execute_script("arguments[0].click();", above)
			handles = dr.window_handles
			for handle in handles:
				dr.switch_to.window(handle)
				if dr.current_url.find(i) != -1:
					lt_info_handle = handle
					t += 1
			dr.switch_to.window(lt_info_handle)
		else:
			dr.switch_to.window(lt_info_handle)
			dr.get(url1)
		dr.switch_to.window(lt_info_handle)
		url2 = 'https://order-rp.rms.rakuten.co.jp/order-rb-mail/individual-mail-send-template-select-sc/init?orderNumber='+i
		dr.get(url2)
		time.sleep(2)
		above = dr.find_element_by_xpath('//*[@id="templateId3"]')
		dr.execute_script("arguments[0].click();", above)
		time.sleep(2)
		above = dr.find_element_by_xpath("//td[@align='center']/input[@class='submitControlButton']")
		dr.execute_script("arguments[0].click();", above)
		time.sleep(2)
		above = dr.find_element_by_xpath("//td[@align='center']/input[2][@class='doubleSubmitInherit submitControlButton']")
		dr.execute_script("arguments[0].click();", above)		
# print(3)
#		dr.switch_to.window(lt_list_handle)
#	dr.refresh()


#欢迎邮件
#发送欢迎邮件条件：
##################################################################################################
#
#1.没有发送过欢迎信件
#判断流程：
#找到属性为"('span',attrs={"data-original-title": 'サンクスメール'}))"的字段并判断class中是否存在'disabled'字样
#已发送欢迎信的class："rms-icon icon light rms-icon-square-mail-1-mini "
#未发送欢迎信的class："rms-icon icon light rms-icon-square-mail-1-mini disabled"
#
##################################################################################################
#
#2.已填入担当者
#判断流程：
#先进入未发送欢迎信件的订单详情，获取id为"orderDetailsFormPersonInCharge-1"的input框的文本内容，并判断是否为空
#
##################################################################################################
#
#整体流程：
#1、订单列表查找符合条件1的订单id并加入list
#2、查找完毕以后点击进入订单（当前应获取所有句柄，并循环切换句柄对应窗口判断当前url来判断句柄对应的窗口是哪个）
#3、获取句柄之后切换到（lt_info_handle）对应的窗口
#4、判断是否符合条件2
#5、不符合条件2时切换回（lt_list_handle）对应的窗口并重新进行循环
#6、符合条件2进入发信url进行发信
#
##################################################################################################
###		欢迎信件需要重写获取条件过程，句柄操作有bug		###
def send_welcome_message():
	amazon_handle = ''
	lt_list_handle = ''
	lt_info_handle = ''
	handles = dr.window_handles
	for handle in handles:
		dr.switch_to.window(handle)
		curl = dr.current_url
		#判断当前url是否为乐天订单列表url
		if curl.find(lt_list_check_url) != -1:
			lt_list_handle = handle
	dr.switch_to.window(lt_list_handle)
	
	# WebDriverWait(dr,10,0.5,ignored_exceptions=None).until(
	# 		EC.element_to_be_clickable((By.XPATH,'//*[@id="filter1"]/ul/li[4]'))
	# 	)
	time.sleep(2)
	above = dr.find_element_by_xpath('//*[@id="filter1"]/ul/li[4]/a/h5/span')
	dr.execute_script("arguments[0].click();", above)
	# above = "$('div[id=filter1] ul li[4] a span').click()"
	# dr.execute_script(above)
	# above = dr.find_element_by_xpath('//*[@id="filter1"]/ul/li[4]/a')
	# dr.find_element_by_xpath("//*[@id='filter1']/ul/li[4]/a").click()
	# ActionChains(dr).click(above).perform()
	# time.sleep(2)
	# dr.find_element_by_xpath("//li[@data-order-list-filter-tab='発送待ち']").click()
	time.sleep(5)
	source = dr.page_source.encode('GBK', 'ignore')
	strsource = source.decode('GBK','strict')
	soup = BeautifulSoup(strsource,"html.parser")
	print(soup)
	tbodys = soup.find_all('tbody',class_ = 'rms-content-order-list-item rms-list-item rms-item-order-inprogress')
	# tbodys = soup.find_all('tbody',name = 'listBody')
	tot = 0
	ids = []
	# print(tbodys)
	print("开始检查符合发信条件的订单")
	for tbody in tbodys:
		# print("tbody")
		span1 = str(tbody.find('span',attrs={"data-original-title": 'サンクスメール'}))
		print(span1)
	#	print(span)
		span2 = tbody.find('span',attrs={"data-sort-column": 'deliveryDate'})
		print(span2)
		if span1.find('disabled') != -1:
			# print("disabled")
			date = span2.get_text()
			if date != '':
				ids.append(tbody.find('a',target = 'order_detail').get_text())
	t = 1
	if len(ids) == 0:
		print("无欢迎信件可发")
		return

	print(ids)
	for i in reversed(ids):
		print("当前正在发送欢迎邮件，订单号为%s\n"%i)
		url1 = 'https://order-rp.rms.rakuten.co.jp/order-rb/individual-order-detail-sc/init?orderNumber='+i
		if t == 1:
			dr.switch_to.window(lt_list_handle)
			time.sleep(2)
			above = dr.find_element_by_link_text(i)
			dr.execute_script("arguments[0].click();", above)
			handles = dr.window_handles
			for handle in handles:
				dr.switch_to.window(handle)
				if dr.current_url.find(i) != -1:
					print(handle)
					print(dr.current_url)
					lt_info_handle = handle
					t += 1
			dr.switch_to.window(lt_info_handle)
		else:
			dr.switch_to.window(lt_info_handle)
			dr.get(url1)
		 
		time.sleep(2)
		dr.switch_to.window(lt_info_handle)
		print(dr.current_url)
		detail_of_person = dr.find_element_by_id("orderDetailsFormPersonInCharge-1").get_attribute("value")
		print(detail_of_person)
		if detail_of_person == '' or detail_of_person is None:
			print("订单号为：{}，未填入担当者，不发送欢迎信件\n".format(id))
			listbox_output.insert(END,"订单号为：{0}，未填入担当者，不发送欢迎信件\n".format(id))
#			dr.switch_to.window(lt_list_handle)
			continue
		else:
			url2 = 'https://order-rp.rms.rakuten.co.jp/order-rb-mail/individual-mail-send-template-select-sc/init?orderNumber='+i
			dr.get(url2)
			time.sleep(2)
			above = dr.find_element_by_xpath("//td[@align='center']/input[@class='submitControlButton']")
			dr.execute_script("arguments[0].click();", above)
			time.sleep(2)
			above = dr.find_element_by_xpath("//td[@align='center']/input[2][@class='doubleSubmitInherit submitControlButton']")
			dr.execute_script("arguments[0].click();", above)	
		#		print(3)
	#		dr.switch_to.window(lt_list_handle)
#	dr.refresh()

#取得乐天网址数据
def get_data():
	amazon_handle = ''
	lt_list_handle = ''
	lt_info_handle = ''
	handles = dr.window_handles
	for handle in handles:
		dr.switch_to.window(handle)
		curl = dr.current_url
		if curl.find(amazon_url) != -1:
			amazon_handle = handle
		elif curl.find(lt_list_check_url) != -1:
			lt_list_handle = handle
		elif curl.find(lt_info_url) != -1:
			lt_info_handle = handle
	dr.switch_to.window(lt_info_handle)
	info_id = dr.find_element_by_xpath("//*[@id='layoutContent']/form/div/div/div/div[2]/div/div[2]/ul[1]/li[1]/a").text
	#收件人地址
	address_info = dr.find_element_by_xpath("//*[@id='rms-content-order-details-block-destination-1-1-options']/div[1]/div[3]/span[1]").text
	#收件人电话
	info_phone = dr.find_element_by_xpath("//*[@id='rms-content-order-details-block-destination-1-1-options']/div[1]/div[3]/span[2]").text
	#收件人姓名   抓取（送付先情报）
	info_name = dr.find_element_by_xpath("//*[@id='rms-content-order-details-block-destination-1-1-options']/div[1]/div[2]/div[2]/span[2]").text
	# //*[@id="rms-content-order-details-block-destination-1-1-options"]/div[1]/div[2]/div[2]/span[2]
	
	info_address = address_info[10:]
	info_zipcode1 = address_info[1:4]
	info_zipcode2 = address_info[5:9]
	info_province = ''
	info_address1 = ''
	info_address2 = ''
	province_location = info_address.find('県')
	
	if province_location == 3:
		info_province = info_address[:4]
		info_address = info_address[4:]
	else:
		info_province = info_address[:3]
		info_address = info_address[3:]
	lenth = len(info_address)
#	listbox_output.insert(END, "当前地址长度为%d\n"%lenth)
	if lenth > 16:
		pass
#		listbox_output.insert(END, "地址长度超过16位" + '\n')
#		info_address2 = address_info[lenth-1]
		
	listbox_output.insert(END, info_id + '\n')
	listbox_output.insert(END, info_name + '\n')
	listbox_output.insert(END, info_zipcode1+'-'+info_zipcode2 + '\n')
	listbox_output.insert(END, info_phone + '\n')
	listbox_output.insert(END, info_province + '\n')
	listbox_output.insert(END, info_address + '\n')

#管理物流信息
def deliver_Manage():
	try:
		#句柄判断
		amazon_handle = ''
		lt_list_handle = ''
		lt_info_handle = ''
		handles = dr.window_handles
		for handle in handles:
			dr.switch_to.window(handle)
			curl = dr.current_url
			#判断当前url是否为亚马逊url
			if curl.find(amazon_check_url) != -1:
				amazon_handle = handle
			#判断当前url是否为乐天订单列表url
			elif curl.find(lt_list_check_url) != -1:
				lt_list_handle = handle
			#判断当前url是否为乐天订单详情url
			elif curl.find(lt_info_url) != -1:
				lt_info_handle = handle
		#切换到乐天订单列表详情
		dr.switch_to.window(lt_list_handle)
		#利用重新登录网址达成刷新目的(直接刷新窗口会有刷不出来新单的情况)
		dr.get(lt_list_url)
		 
		time.sleep(5)
		#切换到待发送新建订单的窗口
		above = dr.find_element_by_xpath('//*[@id="filter1"]/ul/li[4]/a/h5/span')
		dr.execute_script("arguments[0].click();", above)
		# dr.find_element_by_xpath("//*[@id='filter1']/ul/li[4]/a").click()
		# print("ActionChains")
		# above = dr.find_element_by_xpath("//li[@data-order-list-filter-tab='発送待ち' and @data-order-progress='300']/a/span")
		# ActionChains(dr).move_to_element(above).click(above).perform()
		# print("drfindxpath")
		# above.send_keys(Keys.ENTER)
		# print("500")
		# dr.execute_script("window.scrollTo(0, document.body.scrollHeight);")
		# time.sleep(5)
		#解码网页
		source = dr.page_source.encode('GBK', 'ignore')
		strsource = source.decode('GBK','strict')
		soup = BeautifulSoup(strsource,"html.parser")
		Select(dr.find_element_by_xpath('//*[@id="ddlDisplay"]')).select_by_value("500")
		time.sleep(3)
		above = dr.find_element_by_xpath('//*[@id="filter1"]/ul/li[4]/a/h5/span')
		dr.execute_script("arguments[0].click();", above)
		# dr.find_element_by_xpath("//*[@id='filter1']/ul/li[4]/a").click()
		# dr.find_element_by_xpath("//*[@id='filter1']/ul/li[4]/a").click()
		# dr.find_element_by_xpath("//li[@data-order-list-filter-tab='発送待ち']/a").click()
		#查询物流信息
		print('开始查询物流信息')
		check_Deliver_info()
		 
		#发信完毕切换到订单列表的窗口
		dr.switch_to.window(lt_list_handle)
		#利用重新登录网址达成刷新目的(直接刷新窗口会有刷不出来新单的情况)
		dr.get(lt_list_url)
		#点击到待发送信件的订单窗口
		# dr.find_element_by_xpath("//li[@data-order-list-filter-tab='発送待ち']/a/h5/span").send_keys(Keys.ENTER)
	except BaseException as e:
		print(e)
		if str(e).find("no such window: window was already closed"):
			ha = dr.window_handles
			dr.switch_to.window(lt_list_handle)	

#检查发货进度
#发货进度判断条件：
##################################################################################################
#Amazon改了新UI以后物流信息难判断，暂时总结规律
#
#
#1.物流进程到达第二个判断框并出现勾，但是这只能初步判断可能出现物流信息，暂不用此方法判断
#
#2.出现“さらに表示”字样，则说明有详细物流信息，可以点击该字样或者伝票番号来获取详细信息
#
#3.有一种情况不出现“さらに表示”字样但是也能获取信息，即出现“再配達リクエストを受け付けました”字样，疑似出现配送问题，此时最好点击伝票番号来获取详细信息
#
#
##################################################################################################
def check_Deliver_info():
	try:
		amazon_handle = ''
		lt_list_handle = ''
		lt_info_handle = ''
#		print("check_Deliver_info")
		#检查句柄并确认对应网址的标签页的句柄
		handles = dr.window_handles
		for handle in handles:
			dr.switch_to.window(handle)
			curl = dr.current_url
			#判断当前url是否为亚马逊url
			if curl.find(amazon_check_url) != -1:
				amazon_handle = handle
			#判断当前url是否为乐天订单列表url
			elif curl.find(lt_list_check_url) != -1:
				lt_list_handle = handle
			#判断当前url是否为乐天订单详情url
			elif curl.find(lt_info_url) != -1:
				lt_info_handle = handle
				
		#切换到订单列表窗口
		dr.switch_to.window(lt_list_handle)
#		#利用重新登录网址达成刷新目的(直接刷新窗口会有刷不出来新单的情况)
		dr.get(lt_list_url)
		time.sleep(3)
#		#切换到待发送新建的订单列表窗口
		above = dr.find_element_by_xpath('//*[@id="filter1"]/ul/li[4]/a/h5/span')
		dr.execute_script("arguments[0].click();", above)
		# dr.find_element_by_xpath('//*[@id="filter1"]/ul/li[4]/a/h5/span').click()
		# dr.find_element_by_xpath("//li[@data-order-list-filter-tab='発送待ち']/a").send_keys(Keys.ENTER)
		# dr.find_element_by_xpath("//li[@data-order-list-filter-tab='発送待ち']").click()
		# dr.find_element_by_xpath("//li[@data-order-list-filter-tab='発送待ち' and @class = 'filtertab']").click()
		 
		time.sleep(2)
		#对网页进行解码
		source = dr.page_source.encode('GBK', 'ignore')
		strsource = source.decode('GBK','strict')
		soup = BeautifulSoup(strsource,"html.parser")
		tbodys = soup.find_all('tbody',class_ = 'rms-content-order-list-item rms-list-item rms-item-order-inprogress')
		
		#定义list类型的ids列表，储存符合条件的id
		ids = []
		for tbody in tbodys:
			#获取订单生成时间
			order_time = str(tbody.find("span",attrs = {'class','rms-content-order-list-item-order-datetime'}).get_text())
			#获取亚马逊订单id
			amazon_id = str(tbody.find("span",attrs = {'class','rms-content-order-list-item-order-memo'}).get_text())
			#如果订单已经进货并生成亚马逊订单，就不会为空
			if amazon_id != '':
				#裁剪出生成时间的年月日
				day = order_time[0:10]
				#获取当前时间的年月日
				now_day = time.strftime("%Y-%m-%d")
				#如果是一天以前的订单，则查询物流进度，加入id列表
				if now_day > day:
#					print(tbody.find('a',target = 'order_detail').get_text())
					ids.append(tbody.find('a',target = 'order_detail').get_text())
		print(ids)
		
		t = 1
#		print("ids:\n",ids,'\n')
		for i in reversed(ids):
			url1 = 'https://order-rp.rms.rakuten.co.jp/order-rb/individual-order-detail-sc/init?orderNumber='+i
			#如果第一次则点击订单列表的订单生成新的窗口
			if lt_info_handle == '':
				print('新建页面')
				# dr.find_element_by_link_text(i).click()
				newwindow = 'window.open("%s")' % url1
				dr.execute_script(newwindow)
				#生成新窗口以后获取新窗口的句柄
				handles = dr.window_handles
				for handle in handles:
					dr.switch_to.window(handle)
					if dr.current_url.find(i) != -1:
						lt_info_handle = handle
						t += 1
						print("t = ",t,"\n")
				dr.switch_to.window(lt_info_handle)
				continue
			#如果不是第一次则切换到用以生成订单详情的窗口
			else:
				dr.switch_to.window(lt_info_handle)
				dr.get(url1)
#			dr.switch_to.window(lt_info_handle)
			 
			time.sleep(2)
			#如果抓取到的物流号长度不为零，即已经查询过物流信息，切换到列表进行下一次查询
			deliver_id = dr.find_element_by_id("rms-content-order-details-block-destination-1-1-options-group-0-parcel-number").get_attribute("value")
			# //*[@id="rms-content-order-details-block-destination-1-1-options-group-0-parcel-number"]
			if len(deliver_id) > 0:
				dr.switch_to.window(lt_list_handle)
			#如果抓取到物流号未填写，则进行查询
			elif len(dr.find_element_by_id("rms-content-order-details-block-destination-1-1-options-group-0-parcel-number").get_attribute("value")) == 0:
				#获取亚马逊订单号
				# amazon_id1 = dr.find_element_by_id("orderDetailsFormMemoTextArea-1").text
				amazon_id1 = dr.find_element_by_xpath("//*[@id='orderDetailsFormMemoTextArea-1']").text
				dr.switch_to.window(amazon_handle)
				print("AMAZON_ID = "+amazon_id1+'\n')
		#		getsecondurl()
				 
				#进入搜索页面
				search_url = "https://www.amazon.co.jp/gp/your-account/order-history/ref=oh_aui_search?opt=ab&__mk_ja_JP=%E3%82%AB%E3%82%BF%E3%82%AB%E3%83%8A&search="+amazon_id1
				# https://www.amazon.co.jp/progress-tracker/package/ref=ppx_yo_dt_b_track_package?_encoding=UTF8&itemId=mipgttnrnlkrt&orderId=249-9322521-6796603&packageIndex=0&shipmentId=D0whGsDfB&vt=
				# 249-9322521-6796603 
				# dr.get(search_url)
				#进入查询页面（因为直接进入查询页面也会被弹到搜索页面，所以循序进行）
				info_url = "https://www.amazon.co.jp/progress-tracker/package/ref=ppx_yo_dt_b_track_package?_encoding=UTF8&itemId=mipgttnrnlkrt&orderId="+amazon_id1+"&&packageIndex=0&shipmentId=D0whGsDfB&vt="
				print("Info_url : ",info_url,"\n")
				dr.get(info_url)
				#点击配送状况确认
				# dr.find_element_by_link_text("配送状況を確認").click()
			
				# WebDriverWait(dr,15).until(EC.element_to_be_clickable((By.XPATH ,"//*[@id='a-autoid-1']/span")))
				print("点击表示按钮\n")
				print(dr.current_url)
				# dr.find_element_by_xpath("//*[@id='deliveredAddress']/div/span/a").click()
				# dr.find_element_by_xpath("//*[@id='a-autoid-1']/span").send_keys(Keys.ENTER)
				
				# dr.get(info_url)
				sign_url = "https://www.amazon.co.jp/ap/signin?_"
				if dr.current_url.find(sign_url) != -1:
					print("sign\n")
					dr.find_element_by_id("ap_password").send_keys(amazon_pwd)
					dr.find_element_by_id("signInSubmit").click()
					dr.get(search_url)
					dr.find_element_by_link_text("配送状況を確認").click()
				print("Amazon_id = %s\nOrderid = %s\n"%(amazon_id1,i))
				if dr.current_url.find(info_url) == -1:
					print("break\n")
					continue
				#查询配送详情，如果内容匹配为“配送情報は準備が整いしだい更新されます”则说明没有物流信息
				# deliver_info = dr.find_element_by_xpath("//*[@id='a-page']/div[2]/div[6]/div[2]").text
				#新UI判断
				#解码网页
				 
				print("开始解析网页\n")
				source = dr.page_source.encode('GBK', 'ignore')
				strsource = source.decode('GBK','strict')
				soup = BeautifulSoup(strsource,"html.parser")
#				print(str(soup))
				# print("text = ",dr.find_element_by_xpath("//*[@id='deliveredAddress-container'']").text)
				print("さらに表示 is ",str(soup).find("さらに表示"),"\n")
				print("再配達リクエストを受け付けました is ",str(soup).find("再配達リクエストを受け付けました"),"\n")
#				while(1):
#					a = input()
#					print(a)
#					if a == "1":
#						break
				print("さらに表示 = ",str(soup).find("さらに表示"),'\n')
				print("再配達リクエストを受け付けました = ",str(soup).find("再配達リクエストを受け付けました"),'\n')
				print("配達完了 = ",str(soup).find("配達完了"),'\n')
				print("ご不在のため配達できませんでした = ",str(soup).find("ご不在のため配達できませんでした"),'\n')
				print("明日到着予定 = ",str(soup).find("明日到着予定"),'\n')
				if not (str(soup).find("配達完了") == -1):
					
					# dr.find_element_by_xpath("//*[@id='carrierRelatedInfo-container']/div/span/a").click()	
					print("有物流信息，获取中...")	
					# above = dr.find_element_by_xpath("a[@class=a-link-normal tracking-events-modal-trigger tracker-seeDetailsLink]")
					above = dr.find_element_by_link_text('さらに表示')
					dr.execute_script("arguments[0].click();", above)
					# if not (str(soup).find("本日到着予定") == -1 and str(soup).find('明日到着予定') == -1 and str(soup).find('にお届け') == -1) :
					# 	# にお届け済み
					# 	dr.find_element_by_xpath('//*[@id="progressTracker-container"]/div[2]/div/div[2]/div/div/div/div[2]/span/a').send_keys(Keys.ENTER)
					# else:
					# 	dr.find_element_by_xpath("//*[@id='deliveredAddress']/div/span/a").send_keys(Keys.ENTER)
					 
					time.sleep(2)
					source = dr.page_source.encode('GBK', 'ignore')
					strsource = source.decode('GBK','strict')
					soup = BeautifulSoup(strsource,"html.parser")
					# text = soup.find_all('tbody',class_ = 'rms-content-order-list-item rms-list-item rms-item-order-inprogress')
#					source1 = dr.page_source.encode('GBK', 'ignore')
#					strsource1 = source.decode('GBK','strict')
#					soup1 = BeautifulSoup(strsource,"html.parser")
#					deliverInfoDiv = soup1.find('div',id = 'tracking-events-container').get_text().replace('\n','')
#					print(soup.find_all("div",class_ = "a-container"))
					text = dr.find_element_by_xpath("//*[@id='tracking-events-container']/div").text
					# //*[@id="tracking-events-container"]/div
					code = dr.find_element_by_xpath("//*[@id='tracking-events-container']/div/div[2]/h4").text
					code = code[5:]
#					print("deliverInfoDiv = ",deliverInfoDiv,"\n")
					print("text = ",text,"\n")
					print(code)
					handles = dr.window_handles
					info_handle = ''
#					input()
#					print(url1,'\n')
					#获取完信息后切换到乐天订单详情
					for handle in handles:
						dr.switch_to.window(handle)
#						print(dr.current_url)
						if dr.current_url.find(i)!=-1:
#							print("now is current\n")
							info_handle = handle
					dr.switch_to.window(info_handle)
					#清除框内内容并输入
					dr.find_element_by_id("orderDetailsFormEmailStatement-1").clear()
					dr.find_element_by_id("orderDetailsFormEmailStatement-1").send_keys("配送状況の詳細"+Keys.ENTER)
					dr.find_element_by_id("orderDetailsFormEmailStatement-1").send_keys(text+Keys.ENTER)
#					dr.find_element_by_id("orderDetailsFormEmailStatement").send_keys(text2+Keys.ENTER)
#					dr.find_element_by_id("orderDetailsFormEmailStatement").send_keys(text3+Keys.ENTER)
					 
					#匹配到亚马逊物流时点击亚马逊物流按钮
					if text.find("ヤマト") == -1:
						dr.find_element_by_xpath("//*[@id='rms-content-order-details-block-destination-1-1-options-group-0-delivery-company']/option[2]").click()
					# loc = text3.find("番号")
					# code = text3[loc+3:]
					# loc = text3.find("（")
					# code = code.split(" ")[0]
					print(code,'\n')
					#裁剪出物流编号
					code = code.split(':')[1]
					listbox_output.insert(END,'订单号为  %s  物流编号为  %s  \n'%(i,code))
					#输出物流ID
					print("已输出物流ID")
					dr.find_element_by_id("rms-content-order-details-block-destination-1-1-options-group-0-parcel-number").clear()
					dr.find_element_by_id("rms-content-order-details-block-destination-1-1-options-group-0-parcel-number").send_keys(code)
					 
					print("已点击保存按钮")
					#保存乐天信息
					time.sleep(2)
					# dr.find_element_by_xpath("/html/body").send_keys(Keys.CONTROL+Keys.END)
					above = dr.find_element_by_xpath('//*[@id="btnSave"]')
					dr.execute_script("arguments[0].click();", above)
					# dr.find_element_by_xpath("/html/body/div[1]/div[2]/div/div[2]/div/div/form/div/div/div/div[2]/div/div/div[4]/button").click()
					time.sleep(2)
	#				dr.switch_to.window(lt_list_handle)
				else:
					listbox_output.insert(END,'订单号为  %s  暂无物流信息\n'%i)
					print("物流未更新信息\n")
					dr.switch_to.window(lt_list_handle)
					continue

	except BaseException as e:
		print(e)
		pass


#--------------------------------Star--------------------------------------------------
#*****************************************************************************
#* AddressProcess(address_str) 地址处理函数                                     * 
#* 输入参数：不限字符个数地址，省份已去掉                                          *  
#* 返回：3行地址变量，同时存在全局变量 address_3line[]中， 对应填入到亚马逊的3行地址 *          *
#* 如第3行地址为空，已填入'(Rakuten BUBU蔵の注文)'                                *
#* Author: lion                                                                *    
#  Date：2020-Oct-9                                                            *                               *
#*******************************************************************************
 #全局变量，调用完地址处理函数后，返回3行地址，同时也存在这里
def Addressprocess(address_str):
    global address_3line
    address_3line.clear()                   #清空
    
    if len(address_str) <= 16:              #小于等于16个字符，    
        address_3line.append(address_str)   #第1行不处理
        address_3line.append('')            #第2行填''    
        address_3line.append('(Rakuten BUBU蔵の注文)') #第3行加上广告词
        return address_3line
    else:                                    #大于16字符，
        with open('addressbook.txt','a',encoding='utf-8') as fp: #收集地址数据，
            fp.write(address_str+'\n')
        addresswindow=tkinter.Tk()                               #打开文本窗口处理      
        addresswindow.title('乐天地址处理')
        addresswindow.geometry('500x300')
        
        #统计窗口，显示统计字数 （下一步要实时统计）
        addresslistbox=tkinter.Listbox(addresswindow, width=40)
        addresslistbox.place(x=10, y=5, width=40, height=100)
        
        #文本窗口，地址参数载入
        txtContent=tkinter.scrolledtext.ScrolledText(addresswindow, wrap=tkinter.WORD)
        #txtContent.pack(fill=tkinter.BOTH, expand=tkinter.YES)  
        txtContent.place(x=60, y=5, width=400, height=100)
        
        def Addressload(): #把地址装人文本窗口
            txtContent.delete(0.0, tkinter.END)             #清空文本窗口
            txtContent.insert(tkinter.INSERT, address_str)  #载入地址到文本窗口重排
        def Statistics(): #统计每行字数
            str=txtContent.get(0.0, tkinter.END)    #读取的地址文本 
            line=str.split('\n')                    #以换行分割                    
            addresslistbox.delete(0, END)        #清空字数统计窗口
            for i in range(len(line)-1, -1, -1): #把统计好的字数压进统计窗口，最后的行数先进。
                addresslistbox.insert(0, len(line[i]))
        Addressload() #把地址装人文本窗口
        Statistics()  #统计每行字数

        #统计，取消，确认 3个按键处理
        def Statistics_button():    #统计按键
            Statistics()            #统计每行字数
        buttonStatistics=tkinter.Button(addresswindow, text='统计', width=50, command=Statistics_button)
        buttonStatistics.place(x=20, y=180, width=50, height=20)

        def Cancel(): #取消=重新载入数据
            Addressload()   #把地址装人文本窗口
            Statistics()    #统计每行字数
        buttonCancel=tkinter.Button(addresswindow, text='取消', width=50, command=Cancel)
        buttonCancel.place(x=120, y=180, width=50, height=20)

        def Comfirm(): #确认按键，地址的3行数据存到全局变量中
            global address_3line                    #定义address_3line全局变量
            Statistics()                            #统计
            str=txtContent.get(0.0, tkinter.END)    #读出文本窗口排版
            
            address_3line = str.split('\n')         #
            for j in range(0, 3, 1):  #检查地址行字数，>16个字符重新排
                if len(address_3line[j]) > 16:
                    tkinter.messagebox.showerror(title='地址错误', message='地址行字数多于16个\n请重排')
                    return
            
            n = len(address_3line)  #检查地址行数，多于4行的弹出错误窗口要求重排，第3行如空，填入(Rakuten BUBU蔵の注文)
            if n == 2:
                address_3line.append('(Rakuten BUBU蔵の注文)') 
            elif n == 3:
                address_3line[2] = '(Rakuten BUBU蔵の注文)'
            elif n == 4:
                del address_3line[3]
            else:
                tkinter.messagebox.showerror(title='地址错误', message='地址行数多于3行\n请重排')
                return   
            
            #print(address_3line)       #调试时打开这句，关掉下面两句
            addresswindow.quit()        #确认成功，退出地址处理窗口
            addresswindow.destroy()     #关掉窗口
        buttonComfirm=tkinter.Button(addresswindow, text='确认', width=50, command=Comfirm)
        buttonComfirm.place(x=220, y=180, width=50, height=20)
        print(address_3line)
        addresswindow.mainloop()        #弹出地址处理窗口后循环等待，通过"确认"按键关闭跳出
		
        return address_3line
    #----------------------------地址处理函数 -----END-----  
	
#操作两站
def lt_ama():
	try:
		amazon_handle = ''
		lt_list_handle = ''
		lt_info_handle = ''
		handles = dr.window_handles
		for handle in handles:
			dr.switch_to.window(handle)
			curl = dr.current_url
			if curl.find(amazon_check_url) != -1:
				amazon_handle = handle
			elif curl.find(lt_list_check_url) != -1:
				lt_list_handle = handle
			elif curl.find(lt_info_url) != -1:
				lt_info_handle = handle
		dr.switch_to.window(lt_info_handle)
		# source = dr.page_source.encode('GBK', 'ignore')
		# strsource = source.decode('GBK','strict')
		# soup = BeautifulSoup(strsource,"html.parser")
		# data_id = soup.find('a', class_ = 'rms-status-order-nr')
		# info_id = data_id.get_text()
		info_id = dr.find_element_by_xpath("//*[@id='order-details-1']/div/div[6]/div[4]/div/div[1]/div[2]/table/tbody/tr[1]/td[2]").text
		address_info = dr.find_element_by_xpath("//*[@id='rms-content-order-details-block-destination-1-1-options']/div[1]/div[3]/span[1]").text
		info_phone = dr.find_element_by_xpath('//*[@id="rms-content-order-details-block-destination-1-1-options"]/div[1]/div[3]/span[2]').text
		info_name = dr.find_element_by_xpath("//*[@id='rms-content-order-details-block-destination-1-1-options']/div[1]/div[2]/div[2]/span[2]").text
		order_infomation_id = info_id
		print("order_info=",order_infomation_id)
		print("Phone : ",info_phone)
		info_address = address_info[10:]
		info_zipcode1 = address_info[1:4]
		info_zipcode2 = address_info[5:9]
		info_province = ''
		info_address1 = ''
		info_address2 = ''
		province_location = info_address.find('県')
		if province_location == 3:
			info_province = info_address[:4]
			info_address = info_address[4:]
		else:
			info_province = info_address[:3]
			info_address = info_address[3:]
		lenth = len(info_address)
	#	listbox_output.insert(END, "当前地址长度为%d\n"%lenth)
		#切换到Amazon			
		dr.switch_to.window(amazon_handle)
		#有可能出现长时间未操作需要输入密码重新登录的情况，
		#解析网页查找有无输入密码选项来判断是否出现此情况，
		#不能用dr....text，会出现无法查找到元素并跳出的状况
		source = dr.page_source.encode('GBK', 'ignore')
		strsource = source.decode('GBK','strict')
		soup = BeautifulSoup(strsource,"html.parser")
		print(info_address)
		print("Shopping cart Current Url is : ",dr.current_url,'\n')
		dr.get("https://www.amazon.co.jp/gp/cart/view.html?ref=nav_cart")
		#  
		WebDriverWait(dr,15,ignored_exceptions=NONE).until(EC.element_to_be_clickable((By.NAME,"proceedToRetailCheckout")))
	#	dr.find_element_by_name("proceedToCheckout").click()
		dr.find_element_by_xpath("//*[@id='sc-buy-box-ptc-button']/span/input").send_keys(Keys.ENTER)
		sign_url = "https://www.amazon.co.jp/ap/signin?_"
		if dr.current_url.find(sign_url) != -1:
			dr.find_element_by_id("ap_password").send_keys(amazon_pwd)
			dr.find_element_by_id("signInSubmit").click()
			dr.get("https://www.amazon.co.jp/gp/cart/view.html?ref=nav_cart")
			dr.find_element_by_xpath("//input[@name = 'proceedToCheckout']").click()
		 
		time.sleep(3)
		dr.find_element_by_xpath('//*[@id="address-ui-widgets-enterAddressFullName"]').send_keys(info_name)
		dr.find_element_by_xpath('//*[@id="address-ui-widgets-enterAddressFullName"]').send_keys("様")
		time.sleep(1)
		dr.find_element_by_xpath('//*[@id="address-ui-widgets-enterAddressPostalCodeOne"]').send_keys(info_zipcode1)
		dr.find_element_by_xpath('//*[@id="address-ui-widgets-enterAddressPostalCodeTwo"]').send_keys(info_zipcode2)
		dr.find_element_by_xpath('//*[@id="address-ui-widgets-enterAddressPhoneNumber"]').send_keys(info_phone)
		# xpath1 = "//option[@value='%s']" % info_province
		# dr.find_element_by_xpath('//*[@id="address-ui-widgets-enterAddressStateOrRegion"]/span/span').find_element_by_xpath(xpath1).click()
		data_address = Addressprocess(info_address)
		time.sleep(1)
		dr.find_element_by_xpath('//*[@id="address-ui-widgets-enterAddressLine3"]').clear()
		dr.find_element_by_xpath('//*[@id="address-ui-widgets-enterAddressLine3"]').send_keys(data_address[2])
		dr.find_element_by_xpath('//*[@id="address-ui-widgets-enterAddressLine1"]').clear()
		dr.find_element_by_xpath('//*[@id="address-ui-widgets-enterAddressLine1"]').send_keys(data_address[0])
		dr.find_element_by_xpath('//*[@id="address-ui-widgets-enterAddressLine2"]').clear()
		dr.find_element_by_xpath('//*[@id="address-ui-widgets-enterAddressLine2"]').send_keys(data_address[1])
		time.sleep(1)
#		if info_address2 != ''
		

#			dr.find_element_by_id("enterAddressAddressLine2").send_keys(info_address2)
		time.sleep(1)
		
		listbox_output.insert(END, "订单编号为：%s\n姓名：%s\n电话：%s\n邮编：%s-%s\n地址：%s\n"%(info_id,info_name,info_phone,info_zipcode1,info_zipcode2,info_address))
		if len(info_address) <= 16:
			pass
		else:
#			listbox_output.insert(END, "订单编号为：%s\n地址为:%s\n"%(info_id,info_address))
			listbox_output.insert(END, "地址长度超过16字符，请处理完后点击确认购物\n")


	except BaseException as e:
		print(e)
		pass

#确认地址
def after_submit_address():
	try:
		#判断句柄
		amazon_handle = ''
		lt_list_handle = ''
		lt_info_handle = ''
		handles = dr.window_handles
		for handle in handles:
			dr.switch_to.window(handle)
			curl = dr.current_url
			if curl.find(amazon_check_url) != -1:
				amazon_handle = handle
			elif curl.find(lt_list_check_url) != -1:
				lt_list_handle = handle
			elif curl.find(lt_info_url) != -1:
				source = dr.page_source.encode('GBK', 'ignore')
				strsource = source.decode('GBK','strict')
				soup = BeautifulSoup(strsource,"html.parser")
				data_id = soup.find('a', class_ = 'rms-status-order-nr')
				info_id = data_id.get_text()
				lt_info_handle = handle
		dr.switch_to.window(amazon_handle)
		dr.find_element_by_tag_name('body').send_keys(Keys.END)
		# 点击确认按钮
		#  
		# print("第一个按钮")
		# dr.find_element_by_xpath('/html/body/div[5]/div[2]/div/div[3]/div[1]/div/div[1]/form/div/span/span/span/input').send_keys(Keys.ENTER)
		# dr.find_element_by_xpath('/html/body/div[5]/div[2]/div/div[3]/div[1]/div/div[1]/form/div/span/span/span/input').click()
		zhusuo = "$('span[class=a-button-inner] input[aria-labelledby=address-ui-widgets-form-submit-button-announce]').click()"
		dr.execute_script(zhusuo)
		time.sleep(3)
		print("第二个")
		 
		print("3")
		dr.find_element_by_xpath('//*[@id="shippingOptionFormId"]/div[3]/div/div/span[1]/span/input').send_keys(Keys.ENTER)
		
		
		 
		time.sleep(2)
		print("信用卡输入")
		if not dr.find_element_by_xpath('/html/body/div[5]/div/div[2]/div[3]/div/div[2]/div[1]/form/div[1]/div/div/div[1]/div[2]/div/div/div/div/div/div[1]/span/div/label/input').is_selected():
			while(1):
				if dr.find_element_by_xpath('/html/body/div[5]/div/div[2]/div[3]/div/div[2]/div[1]/form/div[1]/div/div/div[1]/div[2]/div/div/div/div/div/div[1]/span/div/label/input').is_selected():
					print("信用卡已选中")
					dr.find_element_by_xpath('/html/body/div[5]/div/div[2]/div[3]/div/div[2]/div[1]/form/div[1]/div/div/div[1]/div[2]/div/div/div/div/div/div[2]/div[2]/div/span/div[3]/input').send_keys('4649882195642418')
					time.sleep(1)
					xykbut = "$('span[class=a-button a-button-primary pmts-button-input] span button').click()"
					dr.execute_script(xykbut)
					break
				else:
					print("信用卡未选中")
					time.sleep(2)
					xyk = "$('div:first	[data-a-input-name = ppw-instrumentRowSelection] label input').click()"
					dr.execute_script(xyk)
					continue
		time.sleep(2)
		print("CONTI")
		conti = "$('input:first [class=a-button-input a-button-text]').click()"
		dr.execute_script(conti)
		print('下一步')
		time.sleep(2)
		dr.find_element_by_xpath('/html/body/div[5]/div[1]/div[2]/div[3]/div/div[2]/div[1]/form/div[2]/div/div/div/span/span/input').send_keys(Keys.ENTER)
		# dr.find_element_by_link_text('続行').send_keys(Keys.ENTER)
		
		print("Before Switch Current Url is : ",dr.current_url,'\n')
		time.sleep(5)
		# dr.find_element_by_xpath("/a[@class = 'no-link a-button-text gift-popover-link a-declarative ']").click()
		# gift_button = "$(a[class=no-link a-button-text gift-popover-link a-declarative ]).click()"
		# dr.execute_script(gift_button)
		#模拟鼠标点击礼物订单
		# above = dr.find_element_by_xpath('/html/body/div[5]/div[1]/div[2]/form/div/div/div/div[1]/div[5]/div/div/div/div[3]/div/div/div[2]/div[1]/div/div[2]/div[6]/div[1]/span/span/a')
		# ActionChains(dr).move_to_element(above).click(above).perform()
		# aboves = dr.find_elements_by_xpath('//span[@class="a-button a-button-small a-button-icon"]/a[@class="no-link a-button-text gift-popover-link a-declarative"]')
		aboves = dr.find_elements_by_xpath('//span[@class="a-button a-button-small a-button-icon"]/span/a[@class="no-link a-button-text gift-popover-link a-declarative "]')
		# aboves = dr.find_elements_by_xpath('//*[@id="spc-orders"]/div/div/div[3]/div/div/div[2]/div[1]/div/div[2]/div[6]/div[1]/span/span/a/b')
		# aboves = dr.find_elements_by_partial_link_text
		print(aboves)
		t = 0
		for above in aboves:
			print("礼物设置")
			dr.execute_script("arguments[0].click();", above)
			# above.send_keys(Keys.ENTER)
			time.sleep(2)
			#不显示金额
			print("点击金额显示")
			inMoney = dr.find_elements_by_xpath('/html/body/div[8]/div/div[2]/div/div[1]/ol/li[3]/div[2]/div[2]/div[4]/label/input')[t]
			dr.execute_script("arguments[0].click();", inMoney)
			time.sleep(4)
			print("输入礼物信息")
			textarea = dr.find_elements_by_xpath("//textarea[@name='gift-message-text']")[t]
			print(textarea.is_displayed())
			textarea.clear()
			textarea.send_keys("受注番号 	")
			textarea.send_keys(info_id)
			textarea.send_keys(Keys.ENTER)
			textarea.send_keys("Rakuten BUBU蔵")
			time.sleep(2)
			submit_button = dr.find_elements_by_xpath('//*[@id="a-popover-1"]/div/div[2]/div/div[1]/ol/li[4]/span[1]/span/input')[t]
			dr.execute_script("arguments[0].click();", submit_button)
			t += 1
		# dr.find_element_by_xpath('/html/body/div[5]/div[1]/div[2]/form/div/div/div/div[1]/div[5]/div/div/div/div[3]/div/div/div[2]/div[1]/div/div[2]/div[7]/div[1]/span/span/a').click()
		# dr.switch_to.window(lt_info_handle)

		# time.sleep(2)
		# source = dr.page_source.encode('GBK', 'ignore')
		# strsource = source.decode('GBK','strict')
		# soup = BeautifulSoup(strsource,"html.parser")
		# data_id = soup.find('a', class_ = 'rms-status-order-nr')
		# info_id = data_id.get_text()
		# dr.switch_to.window(amazon_handle)
		# time.sleep(2)
		print("已点开设定")
		
		# WebDriverWait(dr, 20, 0.5).until(EC.presence_of_element_located(above))
		# ActionChains(dr).click(above).perform()
		
		time.sleep(2)
		print("输入信息")
		
		time.sleep(1)
		print('点击不显示金额')
		# dr.find_element_by_xpath("//input[@type = 'checkbox' and @class = 'a-declarative hide-prices-checkbox']").click()
		time.sleep(2)
		# money = '$("label:first[class=hide-prices-amazon-msg a-checkbox multi-line-checkbox] input").click()'
		# money = dr.find_element_by_xpath("//label[@class = 'hide-prices-amazon-msg a-checkbox multi-line-checkbox']/input")
		
		# if not dr.find_element_by_xpath('/html/body/div[8]/div/div[2]/div/div[1]/ol/li[3]/div[2]/div[2]/div[4]/label/input').is_selected():
		# 	dr.execute_script(money)
		# 	# dr.find_element_by_xpath('/html/body/div[8]/div/div[2]/div/div[1]/ol/li[3]/div[2]/div[2]/div[4]/label/input').send_keys(Keys.SPACE)
		# 	# dr.find_element_by_xpath('/html/body/div[8]/div/div[2]/div/div[1]/ol/li[3]/div[2]/div[2]/div[4]/label/input').click()

		# time.sleep(2)
		# # a-declarative hide-prices-checkbox
		# # Select()
		
		# while(1):
		# 	if dr.find_element_by_xpath("//label[@class = 'hide-prices-amazon-msg a-checkbox multi-line-checkbox']/input").is_selected():
		# 		print("显示金额已选中")
		# 		break
		# 	else:
		# 		print("显示金额未选中")
		# 		dr.execute_script(money)
		# 		continue
		# 	print('点击礼物设置保存')
		# # time.sleep(1)
		# if soup.find('giftForm') == -1:
		# 	above = dr.find_element_by_xpath('//*[@id="a-popover-1"]/div/div[2]/div/div[1]/ol/li[4]/span[1]/span/input')
		# 	# ActionChains(dr).click(above).perform()
		# 	dr.execute_script("arguments[0].click();", above)
		# else:
		# 	above = dr.find_element_by_xpath('//*[@id="giftForm"]/div[1]/div[2]/div/span[1]/span/input')
		# 	dr.execute_script("arguments[0].click();", above)
		# save = "$('span[class=a-button a-button-primary a-spacing-mini set-gift-options-button] span input').click()"
		# dr.execute_script(save)
		# dr.find_element_by_xpath('//*[@id="a-popover-1"]/div/div[2]/div/div[1]/ol/li[4]/span[1]/span/input').send_keys(Keys.ENTER)
		# dr.find_element_by_xpath('/html/body/div[8]/div/div[2]/div/div[1]/ol/li[4]/span[1]/span/input').send_keys(Keys.ENTER)
		# "納品書に金額を表示しない"
		# dr.find_element_by_xpath('//input[@value = "ギフトの設定を保存"]').send_keys(Keys.ENTER)
	except BaseException as e:
		print(e)
		if str(e).find("no such window: window was already closed"):
			dr.switch_to.window(lt_list_handle)		
		pass
#确认购买
def submit_order():
	try:
		amazon_handle = ''
		lt_list_handle = ''
		lt_info_handle = ''
		handles = dr.window_handles
		for handle in handles:
			dr.switch_to.window(handle)
			curl = dr.current_url
			if curl.find(amazon_check_url) != -1:
				amazon_handle = handle
			elif curl.find(lt_list_check_url) != -1:
				lt_list_handle = handle
			elif curl.find(lt_info_url) != -1:
				lt_info_handle = handle
		dr.switch_to.window(amazon_handle)
		print("Before Switch Current Url is : ",dr.current_url,'\n')
		#  
		


		dr.find_element_by_xpath('//*[@id="placeYourOrder"]/span/input').send_keys(Keys.ENTER)
		 
		source = dr.page_source.encode('GBK', 'ignore')
		strsource = source.decode('GBK','strict')
		soup = BeautifulSoup(strsource,"html.parser")
		# get_amazon_id = soup.find('span', class_ = 'a-text-bold')
		# amazon_id = str(get_amazon_id.get_text())
		#diliver_Date = dr.find_element_by_xpath("//*[@id='a-page']/div[2]/div[1]/div/div/div/div/div[1]/ul/li/span/b").text
		# diliver_Date_parser = soup.find(attrs={'class':'a-list-item'})
		# diliver_Date = diliver_Date_parser.find("b").get_text()
		
		info_data = dr.current_url
		print('URL：',info_data,type(info_data),'\n')
		# url = str(url)
		info_data = info_data.split('orderId=')[1]
		info_data = info_data.split('&')[0]
		print('Split url',info_data,'\n')
		# dr.find_element_by_link_text('')
		#购买完成点击注文详细
		# dr.find_element_by_xpath('//*[@id="a-page"]/div[3]/div[1]/div/div/div/div/div[1]/div[2]/div/a[1]').send_keys(Keys.ENTER)
		# above = dr.find_element_by_partial_link_text('注文詳細')
		# ActionChains(dr).click(above).perform()
		time.sleep(3)
		info_url = "https://www.amazon.co.jp/gp/your-account/order-details/ref=ppx_yo_dt_b_order_details_o00?ie=UTF8&orderID="+info_data
		dr.get(info_url)
		# above = dr.find_element_by_xpath('//*[@id="a-page"]/div[3]/div[1]/div[1]/div/div/div/div[1]/div[2]/div/span/a')
		# dr.execute_script("arguments[0].click();", above)
		# date_xpath = '//*[@id="delivery-promise-'+url+'#itemGroupID0"]/span[2]/span'
		time.sleep(3)
		# dliy_date = dr.find_element_by_xpath('//*[@id="orderDetails"]/div[5]/div/div[1]/div[1]/div[1]/span').text()
		source = dr.page_source.encode('GBK', 'ignore')
		strsource = source.decode('GBK','strict')
		soup = BeautifulSoup(strsource,"html.parser")
		dili_date = soup.find('span', class_ = 'a-size-medium a-color-base a-text-bold').get_text()
		print('current\t',dr.current_url,'\n')
		# orderid = dr.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div[1]/div/span[2]/bdi').text()
		# dili_date = dr.find_element_by_xpath('/html/body/div[1]/div[2]/div[5]/div/div[1]/div[1]/div[1]/span').text()
		print("dili_date = \n",dili_date,'\n')
		# print("order id = \n",orderid,'\n')
		dr.switch_to.window(lt_info_handle)
#		dr.find_element_by_xpath("//span[@class = 'a-text-bold']").
		# dr.find_element_by_xpath('//*[@id="orderDetailsFormMemoTextArea-1"]').send_keys(amazon_id.strip())
		

		dr.find_element_by_xpath('//*[@id="orderDetailsFormPersonInCharge-1"]').send_keys("SAT")
		# dr.find_element_by_xpath('//*[@id="rms-content-order-details-block-destination-1-1-options-group-0-shipment-date"]').send_keys(time.strftime("%g%m%d"))
		above = dr.find_element_by_xpath('//*[@id="rms-content-order-details-block-destination-1-1-options-group-0"]/div[3]/div[1]/ul/li[1]/span')
		dr.execute_script("arguments[0].click();", above)
		dili_date = dili_date.strip()
		print('dili_date',dili_date)
		# year = dili_date.split("年")[0]
		# print(year)
		# dili_date = dili_date.split("年")[1]
		# print(dili_date)
		month = dili_date.split("月")[0]
		dili_date = dili_date.split("月")[1]
		day = dili_date.split("日")[0]
		print(day)
		if len(month)<2:
			month = '0'+str(month)
		if len(day)<2:
			day = '0'+str(day)
		dr.find_element_by_xpath('//*[@id="orderDetailsFormDeliveryDate-1"]').send_keys("%s-%s-%s" % (2020,month,day))
		dr.find_element_by_xpath('//*[@id="orderDetailsFormMemoTextArea-1"]').send_keys(info_data.strip())
		dr.find_element_by_xpath("/html/body").send_keys(Keys.CONTROL+Keys.END)
		# above = dr.find_element_by_xpath("/html/body/div[1]/div[2]/div/div[2]/div/div/form/div/div/div/div[2]/div/div/div[4]/button")
		above = dr.find_element_by_xpath('//*[@id="btnSave"]')
		dr.execute_script("arguments[0].click();", above)
	except BaseException as e:
		print(e)
		pass

#批量注文确认
def submit_text():
	try:
		#判断句柄
		amazon_handle = ''
		lt_list_handle = ''
		lt_info_handle = ''
		handles = dr.window_handles
		for handle in handles:
			dr.switch_to.window(handle)
			curl = dr.current_url
			if curl.find(amazon_check_url) != -1:
				amazon_handle = handle
			elif curl.find(lt_list_check_url) != -1:
				lt_list_handle = handle
			elif curl.find(lt_info_url) != -1:
				lt_info_handle = handle
		dr.switch_to.window(lt_list_handle)
		print("批量注文开始")
		dr.get("https://order-rp.rms.rakuten.co.jp/order-rb/order-list-sc/init?&SEARCH_MODE=1&ORDER_PROGRESS=100")
		dr.refresh()
		element = WebDriverWait(dr, 10).until(
			EC.presence_of_element_located((By.XPATH, '//*[@id="rms-checkbox-order-header-checkbox"]'))
		)
		#点击全选按钮
		dr.find_element_by_xpath('//*[@id="rms-checkbox-order-header-checkbox"]').click()
		# above = dr.find_element_by_xpath('//*[@id="rms-checkbox-order-header-checkbox"]')
		# ActionChains(dr).click(above).perform()
		time.sleep(1)
		#点击 “一括処理” 按钮触发下拉菜单
		# dr.find_element_by_xpath('//*[@id="rms-content-order-filter-final-order-btn"]').click()
		above = dr.find_element_by_xpath('//*[@id="filter2"]/div[1]/div/a/span')
		ActionChains(dr).move_to_element(above).click(above).perform()
		time.sleep(1)
		
		above = dr.find_element_by_xpath('//*[@id="rms-content-order-filter-final-order-btn"]')
		ActionChains(dr).click(above).perform()
		time.sleep(1)
		above = dr.find_element_by_xpath('//*[@id="orderConfirm"]')
		ActionChains(dr).click(above).perform()
		# dr.find_element_by_xpath('//*[@id="orderConfirm"]').click()	
	#		dr.find_element_by_xpath("/html/body").send_keys(Keys.CONTROL+Keys.END)
			#  
	#		dr.find_element_by_id("allCheckBox").click()
	#		dr.find_element_by_xpath("//div[@id='orderListTable']/thead[1]/tr/th[1]/div/label/input").click()
		#	dr.find_element_by_xpath("//li[@data-order-list-filter-tab='楽天処理中' and @class = 'filtertab']").click()
			
			# dr.find_element_by_xpath("//a[@class = 'rms-collapse-button collapsed' and @aria-controls = 'collapseOrderListFilterBatchProcessing']").click()
			#  
	#		dr.find_element_by_xpath("//div[@class = 'rms-content-order-details-block-form-title']").click()
			#  
			# dr.find_element_by_id("rms-content-order-filter-final-order-btn").click()
			#  
			# dr.find_element_by_id("orderConfirm").click()
		listbox_output.insert(END, "批量注文确认完毕\n")
	except BaseException as e:
		print(e)
		pass

#登陆Amazon.co.jp
def Amazon_login():
	try:
#		dr.refresh()
		
		newwindow = 'window.open("%s")' % amazon_url
		dr.execute_script(newwindow)
		handles = dr.window_handles
		#判断句柄
		handles = dr.window_handles
		for handle in handles:
			dr.switch_to.window(handle)
			curl = dr.current_url
			if curl.find(amazon_check_url) != -1:
				amazon_handle = handle
			elif curl.find(lt_list_check_url) != -1:
				lt_list_handle = handle
			elif curl.find(lt_info_url) != -1:
					lt_info_handle = handle
		 
		dr.switch_to.window(amazon_handle)
		 
#		dr.find_element_by_xpath("//*[@id='nav-link-accountList']").click()
		source = dr.page_source.encode('GBK', 'ignore')
		strsource = source.decode('GBK','strict')
		soup = BeautifulSoup(strsource,"html.parser")
		email = soup.find("sato@bjgbf.com")
		# print('sato is'str(soup.find("satoさん")))
		if str(soup.find("satoさん")) == -1:
			if str(soup).find("sato@bjgbf.com") != -1:
				print("find email account\n")
				dr.find_element_by_id("ap_password").send_keys(amazon_pwd)
				try:
					WebDriverWait(dr,15).until(EC.element_to_be_clickable((By.XPATH ,"//*[@id='signInSubmit']")))
				finally:
					dr.find_element_by_xpath("//*[@id='signInSubmit']").send_keys(Keys.ENTER)
					dr.find_element_by_xpath("//*[@id='signInSubmit']").click()
					# dr.find_element_by_xpath("//*[@id='signInSubmit']").click()
				
			else:
				dr.find_element_by_id("ap_email").send_keys(amazon_account)
				dr.find_element_by_id("continue").click()
				dr.find_element_by_id("ap_password").send_keys(amazon_pwd)
				dr.find_element_by_xpath("//*[@id='signInSubmit']").send_keys(Keys.ENTER)
	except BaseException as e:
		print(e)
		if str(e).find("no such window: window was already closed"):
			ha = dr.window_handles
			dr.switch_to.window(ha[0])
		pass

#清除日志
def cleanlogs():
	listbox_output.delete(1.0,END)

#结束
def endofroot():
	dr.quit()
	root.destroy()

with open("record.txt",'w+') as f:
	records_names = f.readline().split(',')
	f.close()

Button(root, text = "登陆亚马逊", command = Amazon_login).place(x = 25,y = 50, width = 100,height = 20)
Button(root, text = "开始购物", command = lt_ama).place(x = 25,y = 100, width = 100,height = 20)
Button(root, text = "地址确认", command = after_submit_address).place(x = 25,y = 150, width = 100,height = 20)
#Button(root, text = "提取信息", command = get_data).place(x = 425,y = 250, width = 150,height = 20)
Button(root, text = "确认购物", command = submit_order).place(x = 25,y = 200, width = 100,height = 20)
Button(root, text = "自动操作", command = autoctl).place(x = 25,y = 250, width = 100,height = 20)
Button(root, text = "批量注文", command = submit_text).place(x = 25,y = 300, width = 100,height = 20)
Button(root, text = "一键发信", command = send_message).place(x = 25,y = 350, width = 100,height = 20)
Button(root, text = "物流信息", command = deliver_Manage).place(x = 25,y = 400, width = 100,height = 20)
#Button(root, text = "获得数据", command = get_data).place(x = 50,y = 250, width = 325,height = 20)
# Button(root, text = "清除输出", command = cleanlogs).place(x = 50,y = 320, width = 325,height = 20)
# Button(root, text = "退出程序", command = endofroot).place(x = 425,y = 320, width = 325,height = 20)

listbox_output.insert(END,lt_account1+'\n')
listbox_output.insert(END,pwd1+'\n')
listbox_output.insert(END,lt_account2+'\n')
listbox_output.insert(END,pwd2+'\n')


#Canvas(root, width = 250, height = 50).grid(row = 7, column = 1)

t1 = threading.Thread(target = autoctl)
t1.start()
#t2 = threading.Thread(target = Amazon_login)
#t2.start()
root.mainloop()
