from time 								import sleep
from pywinauto.application 				import Application
from pywinauto 							import win32defines
from cred 								import *
from pywinauto.win32functions 			import SetForegroundWindow, ShowWindow
from selenium.webdriver.chrome.options 	import Options
from bs4 								import BeautifulSoup as soup
from selenium 							import webdriver
from selenium.webdriver.support.ui 		import WebDriverWait
from selenium.webdriver.support 		import expected_conditions as EC
from selenium.webdriver.common.by 		import By
from selenium.common.exceptions 		import TimeoutException, NoSuchElementException
from selenium.webdriver 				import ActionChains
from selenium.webdriver.common.keys 	import Keys

import time, sys, pyautogui as pgui, datetime, keyboard, os, clipboard

def check_db_errors(filename):
    with open(filename) as csv_file:
        line_count = 0
        account = {}
        db_name = {}
        for row in csv_file:
            line  = [ x.strip() for x in row.strip().split(',') ]

            if account.get(line[0],-1) != -1 or db_name.get(line[3],-1) != -1:
                print ("Error in DB.txt File ; ACCOUNT NAME / DATABASE NAME Clash")
                os._exit(0)
            account[line[0]] = True
            db_name[line[3]] = True
			
def isVisible(driver,locator,timeout=20,Id=False, Css=False, Xpath=False, Class=False):
	try:
		if Id:
			element = WebDriverWait(driver,timeout).until(EC.visibility_of_element_located((By.ID,locator)))
		elif Css:
			element = WebDriverWait(driver,timeout).until(EC.visibility_of_element_located((By.CSS_SELECTOR,locator)))
		elif Xpath:
			element = WebDriverWait(driver,timeout).until(EC.visibility_of_element_located((By.XPATH,locator)))
		elif Class:
			element = WebDriverWait(driver,timeout).until(EC.visibility_of_element_located((By.CLASS_NAME,locator)))
		return True
	except TimeoutException:
		return False

def init_driver():
	#driver = webdriver.Chrome("C:\\Users\\1023495\\Downloads\\chromedriver_win32\\chromedriver.exe")
	options = Options()
	options.add_argument("--disable-notifications")
	# options.add_argument('headless')
	driver = webdriver.Chrome(chrome_options=options)
	driver.implicitly_wait(30)
	driver.maximize_window()
	return driver

def goto_app(name):
	# app = Application().connect(path="C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe")
	# try:
	app = Application().connect(title_re = ".*"+name.strip()+".*")
	# except Exception as e:
	# 	print( "Error:: goto_app(): ", sys.exc_info()[0])
	# 	os._exit(1)

	print ("IS HERE")
	w = app.top_window()
	#bring window into foreground
	if w.has_style(win32defines.WS_MINIMIZE): # if minimized
	    ShowWindow(w.wrapper_object(), 9) # restore window state
	else:
	    SetForegroundWindow(w.wrapper_object()) #bring to front

def open_case_queue_sf(driver):
	try:
		driver.get("https://jda--vsupport.cs15.my.salesforce.com/")
		driver.find_element_by_id('username').send_keys(salesforce_id)
		driver.find_element_by_id('password').send_keys(salesforce_passw)
		driver.find_element_by_id('Login').click()

		time.sleep(2)
		# case_tab = driver.find_element_by_xpath('//*[@id="oneHeader"]/div[3]/div/div[1]/div[2]/a')
		case_tab = driver.find_element_by_xpath('//*[@id="oneHeader"]/div[3]/one-appnav/div/one-app-nav-bar/nav/ul/li[2]/a/span')

		case_tab.click()
		print ("Clicked Case Tab")

		# ensuring case page loaded
		WebDriverWait(driver,11).until(EC.visibility_of_element_located((By.XPATH,'//*[@id="brandBand_1"]/div/div[1]/div/div/div/div[2]/div/div[1]/div[2]/div[2]/div[1]/div/div/table/thead/tr/th[3]/div/a')))
		# driver.find_element_by_xpath('//*[@id="oneHeader"]/div[3]/div/div[2]/div/div/ul[2]/li[2]/div[2]/button/lightning-primitive-icon/svg').click()
		# while True:
		# 	try:
		# 		element = driver.find_element_by_xpath('//*[@id="oneHeader"]/div[3]/div/div[2]/div/div/ul[2]/li[2]/div[2]/button/lightning-primitive-icon/svg')
		# 	except NoSuchElementException:
		# 		print ("Done Closing")
		# 		break
		# 	element.click()
		return driver
	except Exception as e:
		print (sys.exc_info()[0].__name__, " in Function open_case_queue_sf()")

def doExecution(driver, case):
	time.sleep(2)
	print ("\nCASE ======>>",case)
	time.sleep(2)
	driver.execute_script("return document.getElementsByTagName('html')[0].innerHTML")
	# element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "myDynamicElement")))
	driver.find_element_by_link_text(case).click()
	driver.find_element_by_link_text(case)
		
	time.sleep(3)

	element = driver.find_elements_by_class_name("test-id__field-value")[18]
	action.move_to_element(element).perform()
	# element.click()
	print (element.text)
	print ("	Page loaded... yeahh")
	time.sleep(2)

	driver.execute_script("window.history.go(-1)")  # command enables going back
	
	return driver

def parse_mail_subject(subject):
	"""
		Takes the Subject returned from the case page
		and returns the date and BU number separately 
		in the desired format.
	"""
	if not subject:
		return -1,-1
	line = [s.strip() for s in subject.split('-')]

	bu = line[1][:2].upper() + " " +  line[1][2:].strip()
	
	month = line[2].split()[1].strip()[:3].lower()

	d = {"jan":1, "feb":2, "mar":3, "apr":4, "may":5, "jun":6, "jul":7, "aug":8, "sep":9, "oct":10, "nov":11, "dec":12}
	month = str(d[month])
	day  = line[2].split()[0].strip()
	year = datetime.datetime.now().year

	if len(month)==1:
		month = "0"+month
	if len(day)==1:
		day = "0" + day

	date =  month + "/" + day + "/" + str(year) 
	# date = day+'/'+month+'/'+str(year)   should be month/date/year
	return bu,date

def handling_new_ones(driver, db_file):
	# goto_app("Sales")
	sleep(10)
	s = driver.execute_script("return document.getElementsByTagName('html')[0].innerHTML")
	f = open('case_scrap.html', 'w', encoding='utf-8', errors = 'ignore')
	f.write(s)
	f.close()

	file = 'case_scrap.html'
	f = open(file,'r' ,encoding='utf-8', errors='ignore')
	again = True
	while again:
		again = False
		try:
			raw_html  = f.read()
		except UnicodeDecodeError:
			again = True

	f.close()

	page = soup(raw_html, features = "html.parser")
	cases = []
	table = page.findAll('tr')
	for i in range(1,len(table)):
		row = table[i]
		case = (row.find('th').text.strip())
		cells = row.findAll('td')

		isDBclear = False
		isNew = False
		for cell in cells:
			if not isDBclear and "dataclear" in cell.text.strip().lower():
				isDBclear = True
			if isDBclear and "new" in cell.text.strip().lower():
				isNew = True

			if isDBclear and isNew:
				cases.append(case)
				break

	print ("Printing Cases: ")
	print (cases)
	action = ActionChains(driver)

	
	for case in cases:
		
		print ("\nCASE ======>>",case)
		time.sleep(2)
		# driver.execute_script("return document.getElementsByTagName('html')[0].innerHTML")
		# element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "myDynamicElement")))
		driver.find_element_by_link_text(case).click()
		driver.find_element_by_link_text(case)
		# driver.refresh()
		elements = driver.find_elements_by_class_name("test-id__field-value")
		count = 1
		for element in elements:
			print (count, "	",element.text)
			count += 1
			if element.text == "New":
				print ("	Found new	")
				action.move_to_element(element).perform()
				print(" move performed ")
				time.sleep(2)


				time.sleep(2)
				action.double_click(element).perform()
				time.sleep(4)
				# driver.find_element_by_link_text("New").click()
				# time.sleep(4)
				# driver.find_element_by_link_text("Work In Progress")
				# time.sleep(4)
				# driver.find_element_by_xpath('//*[@id="brandBand_1"]/div/div[1]/div[3]/div[1]/div/div[2]/div[2]/div/div/div[4]/div/div[2]/div/div/button[2]').click()
				# time.sleep(4)
				break
		driver.execute_script("window.history.go(-1)")  # command enables going back
	"""
		element = WebDriverWait(driver,11).until(EC.visibility_of_element_located((By.XPATH,'//*[@id="primaryField"]/span')))
		# driver.refresh()
		element = WebDriverWait(driver, 20).until(EC.visibility_of_all_elements_located((By.CLASS_NAME, "test-id__field-value" )))
		s = driver.execute_script("return document.getElementsByTagName('html')[0].innerHTML")
		print ("test-id__field-value" in s)
		# element = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CLASS_NAME, "test-id__field-value" )))
		element = driver.find_elements_by_class_name('test-id__field-value')[2]
		print ("Status: ",element.text)
		action.move_to_element(element).perform()
		action.double_click(element).perform()
		time.sleep(1)
		print ("Clicked status")


		# clicking clicked new for dropdown
		element = WebDriverWait(driver, 20).until(EC.visibility_of_any_elements_located((By.CLASS_NAME, "select" )))
		element = driver.find_elements_by_class_name('select')[1]
		action.move_to_element(element).perform()
		element.click()
		time.sleep(1)
		print ("Clicked new")


		#Clicking CASE TAB
		case_tab = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="oneHeader"]/div[3]/one-appnav/div/one-app-nav-bar/nav/ul/li[4]/a/span' )))
		case_tab = driver.find_element_by_xpath('//*[@id="oneHeader"]/div[3]/one-appnav/div/one-app-nav-bar/nav/ul/li[4]/a/span')	
		case_tab.click()
		print ("Clicked CASE")
	
		

		# driver = doExecution(driver,i)
	 
		# get account_name, case number, date and BU number.   
		# Account name to be used to extract DB info from file.
		# Rest three info to change local script. 
	
		x = parse_mail_subject(subject)
		time.sleep(1)
		text = change_notepad_mount_clipboard(filename,bu = x[0] , date = x[1], case_num = case)
		clipboard.copy(text)
		goto_app("Remote")
		connectServer(account_name)
	
	"""
	# print (cases)

def image_exists(name, timeout= 7,arg_confidence=0.9):
	"""
		Checks if and image is present in the screen or not
	"""

	b, timedOut = None, False
	start = datetime.datetime.now()
	while not b:
		# CAUTION CAUTION CAUTION
		# Do not make confidence value as 1
		b = pgui.locateCenterOnScreen(name,grayscale = True, confidence=arg_confidence)
		end = datetime.datetime.now()
		if (( end.minute*60 + end.second)-(start.minute*60 + start.second)) >= timeout:
			timedOut = True
			break
	if timedOut and not b:
		print ("Image ", name, " Doesn't Exists.")
		return False
	return True

def find_click_image(name, timeout = 45, arg_clicks=2, arg_interval=0.005, arg_button = 'left', x_off = 0, y_off = 0,arg_confidence=0.8):
	"""
		Find and click the image in the available visible screen.
		Confidence value is approx the % of pixels you want to 
		match in order to say that its your desired one.
		Default is set to 80%
	"""

	# b will be the tuple of coordinates of the image found
	b, timedOut = None, False
	start = datetime.datetime.now()
	
	while not b:
		# Run indefinitely While the image is not located or the timeout is not reached 

		# CAUTION CAUTION CAUTION
		# Do not make confidence value as 1 (max 0.999)
		b = pgui.locateCenterOnScreen(name,grayscale = True, confidence=arg_confidence)

		end = datetime.datetime.now()
		# See if current time is exceeded the timeout range from start
		if (( end.minute*60 + end.second)-(start.minute*60 + start.second)) > timeout:
			timedOut = True
			break

	if timedOut:
		# Print and Exit after Time Out
		print ("Image ",name , "Not Found for clicking: TimedOut ERROR")
		os._exit(1)
	print ("Clicked: ", name)

	# click image 
	pgui.click(x=b[0]+x_off, y=b[1]+y_off, clicks=arg_clicks, interval=arg_interval, button=arg_button)

def find_image(name, timeout = 45,arg_confidence=0.65):
	"""	
		Find The image "name" till the timeout
	"""
	b, timedOut = None, False
	start = datetime.datetime.now()
	ss = start.minute*60 + start.second
	while not b:
		# CAUTION CAUTION CAUTION
		# Do not make confidence value as 1
		b = pgui.locateCenterOnScreen(name,grayscale = True, confidence=arg_confidence)
		end = datetime.datetime.now()
		es = end.minute*60 + end.second

		if (es-ss) > timeout:
			timedOut = True
			break

	if timedOut:
		print ("Image", name ," Not Found : TimedOut ERROR")
		os._exit(1)

def find_any_click( image_array ):
	""" 
		array elements should always be passed in chunck of 3 elements.  
		First element should be image path, 2nd and 3rd should be an 
		integer which will be treated as offsets of clicking
		in x and in y directions respectively. Ex:- ['img', x offset, y offset,   ... ]
	"""
	if not image_array:
		print ("No image given in find_any_click().... code exit")
		os._exit(1)

	while True:
		# Look for the image for timeout=6 seconds and then iterates over to finding the next 
		for i in range(0,len(image_array)-2,3) :
			if image_exists(image_array[i], timeout=6):
				find_click_image(image_array[i], x_off = image_array[i+1], y_off = image_array[i+2] , arg_confidence=0.9)
				return

def open_mssql():
	App = Application(backend="uia")
	app = App.start("mstsc.exe")
	dlg = app.window(title_re = ".*Remote.*")
	dlg.wait('visible')
	dlg.type_keys(rdp_addr)
	dlg.ConnectButton.click()

	find_click_image('img\\rdp_pass.PNG',arg_clicks=1 , arg_confidence=0.9)
	keyboard.write(rdp_pass+'\n',0.05)
	find_click_image('img\\rdp_ok.PNG',arg_clicks=1 , arg_confidence=0.9)
	find_image('img\\rdp_loaded.PNG' , arg_confidence=0.9)

	# dnc("Opening SQL Server Management Studio")

	find_click_image('img\\rdp_mssql_icon.PNG' , arg_confidence=0.9)
	find_image('img\\rdp_sqlserver_loaded.PNG', arg_confidence = 0.8)
	find_click_image('img\\rdp_sqlserver_name.PNG',x_off=200)

	keyboard.write('DLNPTOTDB02.jdadelivers.com')
	keyboard.send('tab,tab')
	keyboard.write('deploy')
	keyboard.send('tab')
	keyboard.write('deploy#1\n')

	find_image('img\\rdp_server_loaded.PNG', arg_confidence=0.9)
	time.sleep(1)
	find_click_image('img\\rdp_server_newquery.PNG', arg_clicks=1)

	time.sleep(2)
	find_image('img\\rdp_server_query_loaded.PNG')
	time.sleep(2)

def connectServer(account_name):	
	def closeCurrentServer():
		goto_app("Remote")  # focussing rdp
		find_click_image('img\\rdp_server_query_loaded.PNG',x_off=125, y_off = 125, arg_confidence = 0.8) # clicking on text area
		keyboard.send('ctrl+F4')                 # closing opened query
		find_image('img\\rdp_query_closed.PNG')

		keyboard.send('ctrl+F4')                 # disconnecting server
		keyboard.send('alt', do_press = True)
		keyboard.send('f,d,enter')
		keyboard.send('alt', do_release = True)
		
		find_image('img\\rdp_server_disconnected.PNG')
		time.sleep(1)	
	closeCurrentServer()	

	def openServer(account_name):
		time.sleep(1)	
		find_click_image('img\\rdp_server_connect.PNG', arg_clicks=1, arg_confidence=0.9)
		find_image('img\\sqlserver_loaded.PNG', arg_confidence = 0.8)

		find_click_image('img\\rdp_server_name_vm.PNG',x_off=200)

		keyboard.write('DLNPTOTDB02.jdadelivers.com')
		keyboard.send('tab,tab')
		keyboard.write('deploy')
		keyboard.send('tab')
		keyboard.write('deploy#1\n')
		find_image('img\\rdp_server_vm.PNG', arg_confidence=0.9)
		find_image('img\\rdp_server_vm.PNG', arg_confidence=0.9)
		time.sleep(1)
		find_click_image('img\\new_query_vm.PNG', arg_clicks=1)

		time.sleep(2)
		find_image('img\\query_loaded.PNG')
		time.sleep(2)

		find_click_image('img\\master_vm.PNG', arg_confidence = 0.8)
		time.sleep(1)
		keyboard.write('frcoco_test_wh\n')
		time.sleep(2)
	openServer()

def executeScript():
	find_click_image('img\\master_vm.PNG', arg_confidence = 0.8)
	time.sleep(1)
	keyboard.write('frcoco_test_wh\n')
	time.sleep(2)
	find_click_image('img\\query_loaded.PNG',x_off=125, y_off = 125, arg_confidence = 0.8)  

	# dnc("Running the Script in Virtual Machine")

	keyboard.send('ctrl+a')
	keyboard.send('ctrl+v')
	time.sleep(3)
	find_click_image('img\\execute.PNG', arg_confidence=0.9)
	find_image('img\\execute.PNG')

	if image_exists('img\\red_msg.PNG'):
		print("Error in executing the script , please check")
		# dnc ("Script Execution Error")
		os._exit(1)
	elif image_exists('img\\no_go.PNG'):
		print ("No go is there")	
		# dnc("Day already posted")
		goto_app("Sales")
		find_image('img\\home_sf_loaded.PNG')
		find_image('img\\home_sf.PNG')
		find_image('img\\home_sf_loaded.PNG')

		# dnc ("Updating Customer: Day is already posted")
		post_update("The day is posted and data will not be cleared. Thank you")
	
		# dnc("Proceeding to monitor Customer Response")
		post_update(check_change(), fromcustom = True)
	else:
		# print ("checking redprairie\n")
		# dnc("Richie doing UI Validation")
		# check_redprairie(x[0],x[1])
		# dnc("Data Cleared")
		goto_app("Sales")
		post_update("Hello Team, The Data for the Business Date: " + x[1] +" has been cleared. Please verify and confirm case closure. Thank you")
		print  ("DATA CLEARED")
		post_update(check_change(), fromcustom = True)
		os._exit(1)

def rdp_execution():
	try:
		"""
			Open RDP if it's Open Already
		"""
		app = scraper.Application().connect(title_re = ".*Remote.*")
		w = app.top_window()
		#bring window into foreground
		if w.has_style(scraper.win32defines.WS_MINIMIZE): # if minimized
		    scraper.ShowWindow(w.wrapper_object(), 9) # restore window state
		else:
		    scraper.SetForegroundWindow(w.wrapper_object()) #bring to front
	except:
		print ("Opening RDP and MS-SQL")
		open_mssql()

	connectServer()
	executeScript()

 # ############## ################ ############### ############### ######### ######### ######### ########### ########### ######### 