#!/usr/bin/env python3
import pandas as pd
import datetime as dt
import inquirer, openpyxl, zoneinfo, re, os, sys
from O365 import Account, MSGraphProtocol, FileSystemTokenBackend
from dateutil.relativedelta import relativedelta, MO, FR, SU

#https://pypi.org/project/O365/#calendar

CLIENT_ID = '73fee778-cbb5-4c82-81cd-13503338d848'
SECRET_ID = 'SECRET_ID'
credentials = (CLIENT_ID, SECRET_ID)

today = dt.datetime.now()
yesterday = today - dt.timedelta(days=1)
start = None 

#working_dir = os.getcwd()
#working_dir = os.path.dirname(os.path.realpath(__file__))

if getattr(sys, 'frozen', False):
    # If the application is run as a bundle, the PyInstaller bootloader
    # extends the sys module by a flag frozen=True and sets the app 
    # path into variable _MEIPASS'.
    working_dir = os.path.dirname(sys.executable)
else:
	working_dir = os.path.dirname(os.path.abspath(__file__))

#BNE = zoneinfo.ZoneInfo("Australia/Brisbane")

def setup_dirs():
	if not os.path.exists(working_dir+"/exported_formatted"):
		os.makedirs(working_dir+"/exported_formatted")

	return
	
def week_select():
	""" Simplify since never use last week.
	choose_week = [
				inquirer.List('week',
							message="Which week?",
							choices=[
								"This week",
								"Last week", 
								],
						),
			]
	answers = inquirer.prompt(choose_week)
	"""
	answers = {"week":"This week"}
	#answers = {"week":"Last week"}	
	#answers = {"week":"ADHOC"}
	if answers['week'] == "Last week":
		end = yesterday + relativedelta(weekday=SU(-1))
		start = end + relativedelta(weekday=MO(-1))
	elif answers['week'] == "This week" and today.weekday() != 0:
		start = yesterday + relativedelta(weekday=MO(-1))
		end = start + relativedelta(weekday=SU(1))
	elif answers['week'] == "This week" and today.weekday() == 0:
		start = today
		end = start + relativedelta(weekday=SU(1))
	elif answers['week'] == "ADHOC":
		start = dt.datetime(2025,3,24) 
		end = start+relativedelta(weekday=SU(1))
	print("Pulling week: "+str(start.date())+"    ---    "+str(end.date()))
	return dt.datetime.combine(start, dt.datetime.min.time()),dt.datetime.combine(end, dt.datetime.max.time()),start

def get_jobs():
	"""
	Opens excel spreadsheet and pulls in Column A (job short code, ex. 'CAP') and B (job numerical code ex. '000000000') for each row with data
	"""
	job_dict={}

	try:
		jobfile = openpyxl.load_workbook(working_dir+'/jobs_py.xlsx')

	except Exception as e:
		raise e
		print("No job file found. Please create jobs_py.xlsx")
		exit()
	else:
		sheet = jobfile.active
		jobs = list(sheet.values)
		jobfile.close()

		for j in jobs:
			job_dict[j[0]]={"code":j[1],"narrative":{}}
		#print(job_dict)

		return job_dict

def add_job(job_string):

	jobfile = openpyxl.load_workbook(working_dir+'/jobs_py.xlsx')
	sheet = jobfile.active

	#reg_ex find three Capital letters from beginning ex. CAP
	job = re.split(job_string, "^[A-Z]{3}", 1)
	job.append(job_string)

	#find first empty row
	#empty_row = sheet.maxrow()
	empty_row = job
	openfile.save('PATH_TO_FILE.xlsx')
	
	return

def get_events(start,end):

	import warnings

	warnings.filterwarnings(
	    "ignore",
	    message="The localize method is no longer necessary, as this time zone supports the fold attribute",
	)
	warnings.filterwarnings(
	    "ignore",
	    message="The zone attribute is specific to pytz's interface; please migrate to a new time zone provider",
	)

	scopes = ['Calendars.Read','Calendars.Read.Shared']
	token_backend = FileSystemTokenBackend(token_path=working_dir, token_filename='o365_token.txt')
	account = Account(credentials, token_backend=token_backend)
	#print("found account")

	# If it's your first login, you will have to visit a website to authenticate and paste the redirected URL in the console. Then your token will be stored.
	# If you already have a valid token stored, then account.is_authenticated is True.

	if not account.is_authenticated:
		#print("Authenticating")
		account.authenticate(scopes=scopes)
	else:
		print('Authenticated!')

	#print("finding calendar")
	schedule = account.schedule()
	calendar = schedule.get_default_calendar()

	q = calendar.new_query('start').greater_equal(start) #pytz error "localize""
	q.new('end').less_equal(end)

	#events = calendar.get_events() 
	events = calendar.get_events(query=q, limit=500, include_recurring=True) 
	return events

def sort_timesheet(events, job_dict):
	"""
	Sorting events into jobs
	"""

	totals_dict={"total_hours":0,"total_appointments":0,"weekdays":{"0":0,"1":0,"2":0,"3":0,"4":0,"5":0,"6":0},"jobs":{}}
	job_list = []

	for j in job_dict:
		job_list.append(j)
		totals_dict["jobs"][j]=0
	job_list.append("SKIP")

	for event in events:	#pytz error 
		#print(event.start, event.subject)
		dayofweek = event.start.weekday()
		event_date = dt.datetime.strftime(event.start,"%d/%m/%Y")
		duration = abs(event.end-event.start).total_seconds() / 3600
		found=False

		if "SKIP" in event.subject:
				continue

		for j in job_dict:
			#print(j,j in event.subject)

			if j in event.subject and event.subject not in job_dict[j]["narrative"]:
				#print("if:adding row to timesheet_list")
				#job_dict[j]["narrative"][event.subject]={"0":"","1":"","2":"","3":"","4":"","5":"","6":""}
				job_dict[j]["narrative"][event.subject]={"0":"","1":"","2":"","3":"","4":"","5":"","6":""}
				job_dict[j]["narrative"][event.subject][str(dayofweek)]=duration
				found=True

			elif j in event.subject and event.subject in job_dict[j]["narrative"]:
				#print("elif:adding hours to existing row")
				if job_dict[j]["narrative"][event.subject][str(dayofweek)] == "":
					job_dict[j]["narrative"][event.subject][str(dayofweek)]=duration
				else:
					job_dict[j]["narrative"][event.subject][str(dayofweek)]=job_dict[j]["narrative"][event.subject][str(dayofweek)]+duration
				found=True

			if found==True:
				totals_dict["total_hours"]=totals_dict["total_hours"]+duration
				totals_dict["total_appointments"]=totals_dict["total_appointments"]+1
				totals_dict["weekdays"][str(dayofweek)]=totals_dict["weekdays"][str(dayofweek)]+duration
				totals_dict["jobs"][j]=totals_dict["jobs"][j]+duration
				break

		if found == False:
			print(event.start, event.subject)
			questions = [
				inquirer.List('job',
							message="Which job is this? ",
							choices=job_list
							),
						]

			answers = inquirer.prompt(questions)
		
			if answers["job"] == "NEW":
				inp_job = input("ENTER CODE, ex: '074971-01 QLD G2 - CAPACITY (5005-141)'")
				#add_job(inp_job)
				answers["job"] = inp_job
			elif answers["job"] == "SKIP":
				continue

			#print("else:adding row to timesheet_list")
			if event.subject not in job_dict[answers["job"]]["narrative"]:
				job_dict[answers["job"]]["narrative"][event.subject]={"0":"","1":"","2":"","3":"","4":"","5":"","6":""}
			if job_dict[answers["job"]]["narrative"][event.subject][str(dayofweek)] == "":
				job_dict[answers["job"]]["narrative"][event.subject][str(dayofweek)]=duration
			else:
				job_dict[answers["job"]]["narrative"][event.subject][str(dayofweek)]=job_dict[answers["job"]]["narrative"][event.subject][str(dayofweek)]+duration
			totals_dict["total_hours"]=totals_dict["total_hours"]+duration
			totals_dict["total_appointments"]=totals_dict["total_appointments"]+1
			totals_dict["weekdays"][str(dayofweek)]=totals_dict["weekdays"][str(dayofweek)]+duration
			if answers["job"] not in totals_dict["jobs"]:
				totals_dict[answers["job"]][j]=0
			else:
				totals_dict["jobs"][answers["job"]]=totals_dict["jobs"][answers["job"]]+duration
			

	print("Total hours: ",totals_dict["total_hours"])
	print("Total appointments:  ",totals_dict["total_appointments"])
	for j in totals_dict["jobs"]:
		print(j + ": " + str(totals_dict["jobs"][j]))
	print("Mon:  ",totals_dict["weekdays"]["0"])
	print("Tue:  ",totals_dict["weekdays"]["1"])
	print("Wed:  ",totals_dict["weekdays"]["2"])
	print("Thu:  ",totals_dict["weekdays"]["3"])
	print("Fri:  ",totals_dict["weekdays"]["4"])
	print("Sat:  ",totals_dict["weekdays"]["5"])
	print("Sun:  ",totals_dict["weekdays"]["6"])

	return job_dict

def dict_to_rows(timesheet_dict):

	timesheet_list=[]
	"""
	timesheet_list=[
		["Timesheet"],
		["Period Start:",dt.datetime.strftime(start,"%d/%m/%Y")],
		["Time Items"],
		["Charge Code","Date","Hours","Hourly Charge Type","Narrative","Ignore?"],
		]
	"""
	"""
	timesheet_list=[
		[],
		[,dt.datetime.strftime(start,"%d/%m/%Y")],
		[],
		[],
		[],
		]
	"""
	for i in timesheet_dict:
		#print(timesheet_dict[i])
		for n in timesheet_dict[i]["narrative"]:
			timesheet_list.append([
				timesheet_dict[i]["code"],
				"Normal Time",
				n,
				timesheet_dict[i]["narrative"][n]["0"],
				timesheet_dict[i]["narrative"][n]["1"],
				timesheet_dict[i]["narrative"][n]["2"],
				timesheet_dict[i]["narrative"][n]["3"],
				timesheet_dict[i]["narrative"][n]["4"],
				timesheet_dict[i]["narrative"][n]["5"],
				timesheet_dict[i]["narrative"][n]["6"],
				0,
				])

	#print(timesheet_list)

	return timesheet_list

def write_to_excel(timesheet_list,monday_date):
	#print("writing to excel")

	timesheet_file = openpyxl.load_workbook(working_dir+'/Import_TS_Calendar_Hourly.xlsx')
	sheet = timesheet_file.active
	sheet['B2'] = monday_date.date()
	
	for i, line in enumerate(timesheet_list):
		for k, val in enumerate(line):
			cell = sheet.cell(row=i+6, column=k+1)
			#print(val, cell.number_format)
			cell.value = val
			if val == "No":
				cell.number_format = u'"Yes";"Yes";"No"'
			elif isinstance(val, float):
				cell.number_format = u'#,##0.00'
			elif val == "":
				cell.number_format = u'#,##0.00'
	
	timesheet_file.save(working_dir+'/exported_formatted/'+dt.datetime.strftime(monday_date,"%Y%m%d")+'_arup_timesheet'+'.xlsx')

def main():

	setup_dirs()
	start,end,monday_date = week_select()
	job_dict=get_jobs()
	events = get_events(start, end)
	timesheet_dict = sort_timesheet(events, job_dict)
	timesheet_list = dict_to_rows(timesheet_dict)
	write_to_excel(timesheet_list,monday_date)
	print("NOTE: method add_job not completed")
	
	if input("Enter 'r' to run again or any other key to exit.. (press Enter to confirm choice)      ->  ")=="r":
		main()
	
main()              
