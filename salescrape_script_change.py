from pywinauto import application
from bs4 import BeautifulSoup as soup
import pyautogui as pgui, datetime, time, clipboard, cred, keyboard, random, os, tkinter as tk


def dnc(message, life=2):
	"""
		Displays a tkinter GUI with message as "message" (argument)
		and then the GUI destroys after life seconds.
		Life's default value = 2 seconds
	"""
	
	root = tk.Tk()
	# get screen width and height
	ws = root.winfo_screenwidth() # width of the screen
	hs = root.winfo_screenheight() # height of the screen

	w, h  = ws//3.5, hs//6
	x, y = w//5, hs-2*h
	root.title("JDA Richi")
	# set the dimensions of the screen and where it is placed
	root.geometry('%dx%d+%d+%d' % (w, h, x, y))
	# root.geometry("700x200")
	root.attributes("-topmost", True)
	root.configure(background = "#4484CE")

	label1 = tk.Label(root, text=message, fg = "white", bg = "#4484CE", font = ("Arial", 18, "italic"))
	# label1.grid(row = 3, column = 1, rowspan = 10, columnspan = 10)
	label1.pack(fill=tk.BOTH, expand=True)

	root.after(life*1000, lambda: root.destroy())
	root.mainloop()	


	"""
	# Uncommenting this Chunk would bring vertical and horizontal 
	# scroll bars but the text will stick to an edge of GUI

	class AutoScrollbar(tk.Scrollbar):
	    # a scrollbar that hides itself if it's not needed.  only
	    # works if you use the grid geometry manager.
	    def set(self, lo, hi):
	        if float(lo) <= 0.0 and float(hi) >= 1.0:
	            # grid_remove is currently missing from Tkinter!
	            self.tk.call("grid", "remove", self)
	        else:
	            self.grid()
	        tk.Scrollbar.set(self, lo, hi)
	    def pack(self, **kw):
	        raise TclError
	    def place(self, **kw):
	        raise TclError

	# create scrolled canvas
	vscrollbar = AutoScrollbar(root)
	vscrollbar.grid(row=0, column=1, sticky="n"+"s")
	hscrollbar = AutoScrollbar(root, orient=tk.HORIZONTAL)
	hscrollbar.grid(row=1, column=0, sticky="e"+"w")

	canvas = tk.Canvas(root,yscrollcommand=vscrollbar.set, xscrollcommand=hscrollbar.set)
	canvas.grid(row=0, column=0, sticky="n"+"s"+"e"+"w")
	vscrollbar.config(command=canvas.yview)
	hscrollbar.config(command=canvas.xview)
	canvas.configure(background = "#6B7A8F")

	# make the canvas expandable
	root.grid_rowconfigure(0, weight=1)
	root.grid_columnconfigure(0, weight=1)
	# create canvas contents

	frame = tk.Frame(canvas)
	frame.rowconfigure(1, weight=1)
	frame.columnconfigure(1, weight=1)
	

	label1 = tk.Label(frame, text=message, fg = "white", bg = "#6B7A8F", font = ("Arial", 20))
	label1.grid(row = 3, column = 1, rowspan = 10, columnspan = 10)
	#label1.place(x = 100, y =200)

	canvas.create_window(0, 0, anchor="ne", window=frame)
	frame.update_idletasks()
	canvas.config(scrollregion=canvas.bbox("all"))
	"""
	return

def change_notepad_mount_clipboard(filename,bu=-1,date=-1, case_num=-1):
	""" 
		filename is the full path of location of the locally saved script which
		will be changed in this function.

		Always increment the suffix of sales case number , now it's 45.
		Else the edited script won'y run in SQL server in Remote Desktop Connection (RDP).
	"""

	f1 = open(filename,'r')
	text = f1.read()
	f1.close()

	if bu != -1 and date != -1:
		text = text.replace("Business Unit Name",bu.strip().split()[1] )
		# text = text.replace("Business Unit Name",  )
		text = text.replace("<Business Date>","02/07/2018")  # mm/dd/yyy
		text  = text.replace("SF<Case Number>","SF"+case_num+"_45")
	else:
		print ("No messages of DATACLEAR request","change_note")
		# Exiting the code
		os._exit(1)
	return text

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

def change_owner():
	"""
		If the customer reply is inferred as a negative one, then this 
		functions is invoked which changes the owner of the case to
		ESO queue.
	"""

	dnc("Transferring to ESO Queue")

	find_click_image('img\\change_owner.PNG')
	find_image('img\\the_new_owner.PNG')
	find_image('img\\search_people.PNG')
	find_click_image('img\\search_people.PNG')
	keyboard.write('ESO')
	find_click_image('img\\eso_core_team.PNG')
	find_click_image('img\\change_owner_submit.PNG')

def close_case():
	"""
		If the customer reply is inferred as a positive one, then this 
		functions is invoked which proceeds to post a resolution of 
		closing the case after changing its status to closed.
	"""
	dnc("Closing the Case")

	find_any_click(['img\\change_status.PNG',0,0,'img\\change_status1.PNG',0,0])
	find_image('img\\change_status_loaded.PNG', arg_confidence=0.9)
	find_click_image('img\\change_status_loaded.PNG', arg_confidence=0.9)
	find_click_image('img\\close_wip.PNG', arg_confidence=0.9, arg_clicks=1)
	find_click_image('img\\close_drop.PNG', arg_confidence=0.9, arg_clicks=1)

	for i in range(3):
		pgui.scroll(-200)
	find_image('img\\category.PNG')
	find_click_image('img\\category.PNG', arg_confidence=0.9)
	keyboard.write('DATA')
	time.sleep(2)
	find_image('img\\data_bhavana.PNG')
	find_click_image('img\\data_bhavana.PNG', arg_confidence = 0.9, arg_clicks = 1)
	pgui.scroll(-150)

	while not image_exists('img\\resolution_left_align.PNG', timeout=3):
		if image_exists('img\\resolution.PNG', timeout=3):
			find_click_image('img\\resolution.PNG', arg_confidence=0.9, arg_clicks=1, y_off = 40)
	
	time.sleep(3)
	keyboard.write('Issue resolved. Customer confirmed to close the case',0.05)
	find_click_image('img\\save.PNG', arg_confidence=0.9, arg_clicks=1)

def save_source(file):
	"""
		I failed to extract page (which comes after login
		authentication and redirection ) source code with 
		bs4, beautifulSoup. This function extracts source
		code in "file" path by inspecting page. Function 
		works only with Google Chrome Browser and only on 
		test.salesforce.com lightining app version.
	"""

	# Following 3 lines can be commented to extract source code of any page
	# given the page of the chrome is in focus among other windows.

	find_image('img\\sanbox_operations.PNG', arg_confidence=0.9)
	find_click_image('img\\cloud.PNG', arg_clicks=1 , arg_confidence=0.9)
	find_click_image('img\\sanbox_operations.PNG', arg_clicks = 1, arg_confidence=0.9)

	pgui.hotkey('ctrl','shift','i')
	time.sleep(2)
	find_click_image('img\\element.PNG' , arg_confidence=0.9)
	time.sleep(1)
	find_click_image('img\\html.PNG', arg_button = 'right',arg_clicks=1 , arg_confidence=0.9)
	time.sleep(1)
	find_click_image('img\\copy.PNG',arg_clicks=1 , arg_confidence=0.9)
	find_click_image('img\\copy_outer.PNG',arg_clicks=1 , arg_confidence=0.9)
	pgui.hotkey('ctrl','shift','i')

	f = open(file,'w',encoding='utf-8', errors='ignore')
	source = clipboard.paste()
	f.write(source)
	f.close()

def reply(customer_reply):
	"""
		Python uses this function to reply to the customers.
		Two array of keywords positives(poss) and negatives(negs)
		are defined. Based on presence of these keyword customer reply's
		sentiment is decided and then a random reply is again
		choosen from two arrays of positive replies and negative
		replies.
	"""
	customer_reply = customer_reply.lower()

	poss = ["close", "closure", "resolved", "thanks!!","thank", "resolve", "closed", "Thank"]
	negs = ["don't", "do not", "no", "can not", "can't", "shouldn't", "not"]

	proceed = False
	for word in poss:
		if word in customer_reply:
			proceed = True

	for word in negs:
		if word in customer_reply:
			proceed = False

	positive_reply = ["Great !, we'll close the case.",	"Cool !, we'll close the ticket","Okay ! case will be closed now", "Good to hear that, we'll close the case"]
	apology_reply = ["Sorry, for the inconvenience caused !!, we'll look into the issue", "Apologies, we'll look into the issue", "Thanks for informing !, we'll look into the issue"]

	"""
		bool good is the decided sentiment i.e., positive or negative
		of the customer_reply and returned for further processing.
	"""
	good = True
	if proceed:
		customer_reply = (random.choice(positive_reply))
	else:
		good = False
		customer_reply = (random.choice(apology_reply)) 
	return customer_reply, good

def post_update(comment, fromcustom = False):
	"""
		Works only if salesforce lightning app version is visible on the screen.
		Clicks on POST and types update and also replies to the customer response.
		Also invokes change_owner() or close_case() according to the positive/negative
		reply from the customer.
	"""
	find_click_image('img\\sanbox_operations.PNG',arg_clicks=1 , arg_confidence=0.9)
	time.sleep(2)
	find_click_image('img\\post.PNG', arg_confidence=0.9)

	find_any_click(['img\\update.PNG',0,0,'img\\update1.PNG',0,-100])
	
	change_own, close_it  = False, False

	if not fromcustom:
		time.sleep(1)
		keyboard.write(comment,0.05)
	else:

		x = reply(comment)
		keyboard.write(x[0])
		if not x[1]:
			dnc("Customer informed: Issue not Resolved")
			change_own = True
		else:
			dnc("Customer confirmed for case closure")
			close_it = True


	time.sleep(3)
	find_click_image('img\\update_share.PNG',arg_clicks=1)
	find_image('img\\update_share.PNG')
	time.sleep(3)
	
	if change_own:
		change_owner()
	if close_it:
		close_case()

	save_source('page0.html')

def get_dataclear_case():
	"""
		Function works only in the page which contains
		case queue. Scrapes code and Find the "dataclear"
		subject.
	"""
	save_source('case_scrap.html')
	file = 'case_scrap.html'
	f = open(file,'r' ,encoding='utf-8', errors='ignore')

	"""
		Reading directly the source code (f.read()) ocassionally gave
		UnicodeDecodeError but not always. So a loop is run continuosly
		till the error doesn't came then breaks.
	"""
	again = True
	while again:
		again = False
		try:
			raw_html  = f.read()
		except UnicodeDecodeError:
			again = True

	f.close()
	page = soup(raw_html, features = "html.parser")
	print (page.title.text)


	# finding the table, iterating each cell in every row(cells) of table
	first = True
	table = page.findAll('tr')
	for i in range(1,len(table)):
		row = table[i]
		case = (row.find('th').text.strip())
		cells = row.findAll('td')

		for cell in cells:
			if "dataclear" in cell.text.strip().lower():
				return case

def get_dataclear_subj():
	"""
		Function works only in the page which appears after
		any case is clicked. Scrapes code and finds the subject
		and returns it.
	"""
	save_source('dataclear_page.html')
	file =  'dataclear_page.html'
	f = open(file,'r',encoding='utf-8', errors='ignore')

	again = True
	while again:
		again = False
		try:
			raw_html  = f.read()
		except UnicodeDecodeError:
			again = True

	f.close()
	page = soup(raw_html, features = "html.parser")

	# Itearting through every header
	for x in page.findAll('h1'):
		ans = x.text.strip()
		if "dataclear" in ans.lower():
			return ans

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
	
	month = line[2][2:].strip()[:3].lower()

	d = {"jan":1, "feb":2, "mar":3, "apr":4, "may":5, "jun":6, "jul":7, "aug":8, "sep":9, "oct":10, "nov":11, "dec":12}
	month = str(d[month])
	day  = line[2][:2]
	year = datetime.datetime.now().year
	date =  str(year) + "-" + month + "-" + day 
	# date = day+'/'+month+'/'+str(year)

	# print (bu,date)
	return bu,date	

def salesforceScrape():

	app = application.Application()
	pgui.hotkey('win','d')

	pgui.hotkey('win','r')
	keyboard.write('chrome about:blank\n')
	time.sleep(2)
	dnc("Checking if Salesforce is Open")
	pgui.hotkey('ctrl','w')

	if image_exists('img\\sanbox_operations.PNG'):
		skip_something = True
		pgui.hotkey('alt','space')
		pgui.hotkey('x')
		print ("\n\nSaelsforce is opened already\n\n")
	else:
		skip_something = False
		print ("\n\nSalesforce is not opened already\n\n")

	if not skip_something:
		# open salesforce in chrome
		pgui.hotkey('win','r')
		keyboard.write('chrome test.salesforce.com\n')
		dnc("Opening Salesforce")

		find_image('img\\page_loaded.PNG')
		pgui.hotkey('alt','space')
		pgui.hotkey('x')
		find_image('img\\test_sf_login_page.PNG')
		find_click_image('img\\log_in_sf.PNG')


		# checks if page is loaded
		find_image('img\\home_sf_loaded.PNG', arg_confidence=0.9)
		find_image('img\\home_sf.PNG' , arg_confidence=0.9)
		find_image('img\\home_sf_loaded.PNG' , arg_confidence=0.9)
		time.sleep(2)

		# goes to case page list
		find_click_image('img\\click_case.PNG' , arg_confidence=0.9)


	find_image('img\\case_loaded.PNG' , arg_confidence=0.9)
	find_image('img\\page_loaded.PNG' , arg_confidence=0.9)
	dnc("Scanning the Queue for Dataclears")

	find_click_image('img\\sanbox_operations.PNG',arg_clicks = 1 , arg_confidence=0.8)
	case =  (get_dataclear_case())
	dnc("Dataclear Case Number Retrieved")
	print (case)


	# searchs and goes into the case
	find_image('img\\search_case_more.PNG', arg_confidence=0.8)
	find_click_image('img\\search_case_more.PNG' , arg_confidence=0.8, arg_clicks = 1)
	keyboard.write(case,0.03)
	time.sleep(3)
	find_click_image('img\\dataclear_case.PNG' , arg_confidence=0.9)
	
	dnc ("Opening First Dataclear Case")
	# changes status to work in progress
	find_image('img\\case_sidebar_account_name.PNG', arg_confidence=0.9)
	# find_click_image('img\\sidebar_case.PNG',arg_clicks=1)
	find_image('img\\case_sidebar_loaded.PNG')	
	
	dnc("Changing Status to Work in progress")

	find_click_image('img\\new.PNG' , arg_confidence=0.9)
	find_click_image('img\\new_drop.PNG', arg_clicks=3 , arg_confidence=0.9)

	while not image_exists('img\\work_in_progress_drop.PNG', timeout = 3):
		find_click_image('img\\new_drop.PNG', arg_clicks=3, timeout = 3 , arg_confidence=0.9)

	find_click_image('img\\work_in_progress_drop.PNG', arg_clicks=1 , arg_confidence=0.9)
	find_click_image('img\\save.PNG', arg_clicks =1 , arg_confidence=0.9)
	find_image('img\\workinprogress.PNG' , arg_confidence=0.9)
	
	dnc("Acknowledging the Customer")

	# posts that the case is acknowledged 
	post_update("Hello Team, This is to acknowledge the case and we are working on your request. Thanks, Jda")
	find_image('img\\update.PNG',  arg_confidence=0.9)
	pgui.hotkey('ctrl','r')

	# checks if page is reloaded
	find_image('img\\home_sf_loaded.PNG' , arg_confidence=0.9)
	find_image('img\\cloud.PNG' , arg_confidence=0.9)
	find_image('img\\home_sf_loaded.PNG' , arg_confidence=0.9)
	time.sleep(2)


	dnc("Marking the First Response to Complete")
	# marks first response
	# find_click_image('img\\milestones.PNG', arg_clicks = 0)
	# pgui.scroll(-200)
	time.sleep(2)
	find_click_image('img\\first_response.PNG',y_off=42)
	time.sleep(4)

	subj = get_dataclear_subj()
	return subj, case

def check_change():
	save_source('page0.html')
	f = open('page0.html','r' ,encoding='utf-8', errors='ignore')
	raw_html  = f.read()
	f.close()
	page = soup(raw_html, features = "html.parser")

	array = [div.find('span') for div in page.findAll("div", {"class": "cuf-feedBodyText forceChatterMessageSegments forceChatterFeedBodyText"})]
	# for i in range(len(array)):
	# 	print (i,array[i].text)
	initial = array[0].text

	last = initial
	while initial == last:
		find_image('img\\sanbox_operations.PNG')
		pgui.hotkey('ctrl','r')
		time.sleep(7)
		save_source('page1.html')
		print ("--------------------")
	
		f = open('page1.html','r' ,encoding='utf-8', errors='ignore')
		raw_html  = f.read()
		f.close()
		page = soup(raw_html, features = "html.parser")

		array = [div.find('span') for div in page.findAll("div", {"class": "cuf-feedBodyText forceChatterMessageSegments forceChatterFeedBodyText"})]
		# for i in range(len(array)):                                       cuf-feedBodyText forceChatterMessageSegments forceChatterFeedBodyText
		# 	print (i,array[i].text)
		last = array[0].text


	return last

def check_redprairie(bu, date):
	app = application.Application()
	pgui.hotkey('win','d')


	dnc("Logging into ESO")
	# open bluecube in internet expxlorer
	pgui.hotkey('win','r')
	keyboard.write('iexplore http://fr-test.jdadelivers.com\n')
	find_click_image('img\\red_prairie_login.PNG', arg_clicks = 1)
	find_click_image('img\\rp.PNG',arg_clicks=1)
	time.sleep(1)
	pgui.hotkey('alt','space')
	pgui.hotkey('x','tab')
	
	time.sleep(2)
	keyboard.write('AAmritanshu',0.05)
	pgui.hotkey('tab')
	keyboard.write('1234567A\n',0.05)
	pgui.hotkey('enter')
	find_image('img\\rp.PNG')
	find_click_image('img\\rp.PNG',arg_clicks=6)
	find_click_image('img\\rp_continue.PNG', arg_confidence=0.7)

	find_image('img\\merchandise_management.PNG')
	find_click_image('img\\workflow_right.PNG', x_off = 220, arg_clicks = 2)

	keyboard.write('upload logs ')
	pgui.hotkey('down','enter')
	find_image('img\\pos_org_unit.PNG')
	find_image('img\\pos_search_result.PNG')
	find_click_image('img\\pos_org_unit.PNG', x_off = 150)
	keyboard.write ( bu.strip().split()[1] )    # BU NUMBER
	find_click_image('img\\pos_first_date.PNG', x_off = 150, arg_clicks=1)
	time.sleep(1)
	keyboard.write('07/02/2018')                # DATE FIRST
	find_click_image('img\\pos_last_date.PNG', x_off = 150, arg_clicks=1, arg_confidence=0.9)
	time.sleep(1)
	keyboard.write('07/02/2018')                # LAST DATE
	find_click_image('img\\pos_view_result.PNG', arg_clicks = 1)
	find_image('img\\pos_business_unit.PNG')

	find_click_image('img\\bus_unit.PNG', arg_confidence=0.8, y_off = 18)
	time.sleep(5)
	find_image('img\\checksum_validation.PNG')
	time.sleep(2)
	dnc("Checking if Data is cleared in UI")
	find_click_image('img\\checksum_validation.PNG')

	if image_exists('img\\nothing_below_host.PNG'):
		dnc("Data Cleared")
		pgui.keyDown('alt')
		pgui.hotkey('tab','tab')
		pgui.keyUp('alt')
		post_update("Hello Team, The Data for the Business Date: " + date +" has been cleared. Please verify and confirm case closure. Thank you")
		print  ("DATA CLEARED")
		# post_update(check_change())
	else:
		# if checksum is not empty
		dnc("Data not Cleared")
		pgui.keyDown('alt')
		pgui.hotkey('tab','tab')
		pgui.keyUp('alt')
		post_update("Dataclear Script ran, but still data Exists, please look into it.")
		change_owner()
		print("DATA NOT CLEARED")
		os._exit(1)

	# if image_exists('img\\no_data.PNG'):   # no data in the checksum validation 
	# 	# ideal thing to happen
	# 	post_update("Hi team, As requested we have deleted the DB. Please check and confirm case closure. Thank You",fromcustom = True)
	# else:
	# 	# unwanted situation
	# 	post_update("Hi team, As requested we have deleted the DB. Please check and confirm case closure. Thank You",fromcustom = True)

def main1(bu,subj):
	filename = "C:\\Users\\1023495\\Downloads\\shell-gss script.txt"
	filename = "C:\\Users\\1023495\\Downloads\\Automation_script_final.txt"

	ret = salesforceScrape()
	
	dnc("Fetching BU name & date")
	print (ret)
	x = parse_mail_subject(ret[0])



	case = ret[1]	
	# print(x[0],x[1])
	bu.append(x[0])
	subj.append(x[1])

	time.sleep(1)
	dnc ("Updating the Script")
	text = change_notepad_mount_clipboard(filename,bu = x[0] , date = x[1], case_num = case)
	clipboard.copy(text)

	# print (text)
	
	app = application.Application()
	pgui.hotkey('win','d')
	app.start("mstsc.exe")

	dnc("Logging to Virtual Machine")

	find_click_image('img\\type_rdp_highlight.PNG' , arg_confidence=0.9)
	keyboard.write('10.0.135.120\n',0.05)

	find_click_image('img\\pass_rdp.PNG',arg_clicks=1 , arg_confidence=0.9)
	keyboard.write(cred.passw+'\n',0.05)

	find_click_image('img\\ok_rdp.PNG',arg_clicks=1 , arg_confidence=0.9)

	find_image('img\\vm_loaded.PNG' , arg_confidence=0.9)

	dnc("Opening SQL Server Management Studio")

	find_click_image('img\\mssql_vm.PNG' , arg_confidence=0.9)

	find_image('img\\sqlserver_loaded.PNG', arg_confidence = 0.8)

	find_click_image('img\\server_name_vm.PNG',x_off=200)

	keyboard.write('DLNPTOTDB02.jdadelivers.com')
	keyboard.send('tab,tab')
	keyboard.write('deploy')
	keyboard.send('tab')
	keyboard.write('deploy#1\n')

	find_image('img\\server_vm.PNG', arg_confidence=0.9)
	time.sleep(1)
	find_click_image('img\\new_query_vm.PNG', arg_clicks=1)

	time.sleep(2)
	find_image('img\\query_loaded.PNG')
	time.sleep(2)

	find_click_image('img\\master_vm.PNG', arg_confidence = 0.8)
	time.sleep(1)
	keyboard.write('frcoco_test_wh\n')
	time.sleep(2)
	find_click_image('img\\query_loaded.PNG',x_off=125, y_off = 125, arg_confidence = 0.8)  

	dnc("Running the Script in Virtual Machine")

	keyboard.send('ctrl+a')
	keyboard.send('ctrl+v')
	time.sleep(3)
	find_click_image('img\\execute.PNG', arg_confidence=0.9)
	find_image('img\\execute.PNG')

	if image_exists('img\\red_msg.PNG'):
		print("Error in executing the script , please check")
		dnc ("Script Execution Error")
		os._exit(1)
	elif image_exists('img\\no_go.PNG'):
		print ("No go is there")	
		dnc("Day already posted")
		time.sleep(1)
		pgui.hotkey('alt','tab')
		find_image('img\\home_sf_loaded.PNG')
		find_image('img\\home_sf.PNG')
		find_image('img\\home_sf_loaded.PNG')

		dnc ("Updating Customer: Day is already posted")
		post_update("The day is posted and data will not be cleared. Thank you")
	
		dnc("Proceeding to monitor Customer Response")
		post_update(check_change(), fromcustom = True)

	else:
		print ("checking redprairie\n")
		dnc("Richi doing UI Validation")
		# check_redprairie(x[0],x[1])
		dnc("Data Cleared")
		pgui.keyDown('alt')
		# pgui.hotkey('tab','tab')                        # uncomment this line and comment below line if check_redprairie function is uncommented	
		pgui.hotkey('tab')
		pgui.keyUp('alt')
		post_update("Hello Team, The Data for the Business Date: " + x[1] +" has been cleared. Please verify and confirm case closure. Thank you")
		print  ("DATA CLEARED")
		post_update(check_change(), fromcustom = True)
		os._exit(1)
		

if __name__ == "__main__":
	subj, bu = [], []
	main1(bu, subj)
