#!/usr/bin/env python
"""
This script downloads data from project management site --teamwork.com-- using it's API, puts this into several pandas dataframes, and generates several views on the data in an Excel output file. 
It focuses on time planned and time logged either per person, per project or per task. 
It can be used to quickly view which teammember is overworked or overbooked, and which tasks are not progressing fast enough. 
At INCAS3 it was common to look at past week's progress and coming week's planning so these have a central role. 
In addition, the total time booked on any task/project is also reported - handy for project progress monitoring.
"""

__authors__   = "Jan Stegenga"
__contact__   = "Jan Stegenga <djanstegenga@incas3.eu>"
__copyright__ = "Stichting INCAS3"
__license__   = "INCAS3 Internal"
__date__      = "06.12.2016"
__version__   = "1.0"
__status__    = "Development"

import matplotlib.pyplot as plt
import numpy as np
import os.path
import datetime
import json
import requests
import pandas as pd
from requests.auth import HTTPBasicAuth

#parameters
url = ""																#company link: "https://<company_name>.teamwork.com/"
key = ""																#generate a key in your profile on teamwork projects
fromtime   = datetime.date.today() - datetime.timedelta( days=7 )		#in our company it was common practice to look at last week and coming week
totime     = datetime.date.today()										#
nexttime   = datetime.date.today() + datetime.timedelta( days=6 )		#
outputfile = "C:\\Software\\python\\teamwork-" + datetime.date.strftime(totime, '%Y%m%d' ) + ".xlsx" 

#closes excel if it is opened
os.system('taskkill /IM excel.exe')

#get id of company
def req_id( url ):		
	print( 'request for companies' )
	payload = { "Content-Type": "application/json"}
	r = requests.get( url + "companies.json", auth=HTTPBasicAuth( key, 'xxx' ), params=payload )
	df = pd.DataFrame( r.json()['companies'] )
	incas3id = df.loc[df.isowner == '1'].id[0]
	return incas3id

#get all people
def req_people( url, company_id ):	
	print( 'request for all people' )
	payload = { "Content-Type": "application/json"}
	r = requests.get( url + "/companies/" + company_id + "/people.json", auth=HTTPBasicAuth( key, 'xxx' ), params=payload )
	#print( r )
	print(  'page ' + r.headers['x-page'] + ' of ' + r.headers['x-pages'] )
	df = pd.DataFrame(r.json()['people'])
	df['short-name']= df['first-name'] + ' ' + df['last-name'].str[0] + '.'
	return df
	
#get all logged time	
def req_logtime( url ):		
	print( 'request for all time logged in certain period' )
	payload = { 'fromdate': datetime.date.strftime(fromtime, '%Y%m%d' ),
				'todate': datetime.date.strftime(totime, '%Y%m%d' ),
				"Content-Type": "application/json"}
	r = requests.get( url + "time_entries.json", auth=HTTPBasicAuth( key, 'xxx' ), params=payload )
	print( r )
	#if the number of items is large, multple pages can be requested. This is in r.headers['x-pages']
	print(  'page ' + r.headers['x-page'] + ' of ' + r.headers['x-pages'] )
	df = pd.DataFrame( r.json()['time-entries'] )
	#print( df.columns )
	df = df[ ['person-last-name', 'project-name', 'todo-item-name', 'hours', 'minutes', 'taskEstimatedTime', 'date'] ]
	df = df.convert_objects(convert_numeric=True)
	df[ 'date2' ] = df['date'].apply(pd.to_datetime)
	df[ 'date2' ] = df.date2.dt.date
	df[ 'frac-time' ] = df.hours + df.minutes // 60 + ( df.minutes%60 )/60
	df.drop( ['hours', 'minutes'], axis=1, inplace=True )
	#dftime  = df
	#grptime = df.groupby( 'person-last-name' )
	return df
		
#get all tasks, updates leave tasks
def req_tasks( url, totime, nexttime ):
	print( 'request for all tasks in certain period' )
	payload = { "Content-Type": "application/json"}
	r = requests.get( url + "tasks.json", auth=HTTPBasicAuth( key, 'xxx' ), params=payload )
	#print( r )
	print(  'page ' + r.headers['x-page'] + ' of ' + r.headers['x-pages'] )
	df = pd.DataFrame(r.json()['todo-items'])
	#print( df.columns )
	df = df[ ['project-name', 'start-date', 'due-date', 'responsible-party-lastname', 'responsible-party-names', 'progress', 'estimated-minutes', 'content', 'recurring', 'completed', 'id', 'todo-list-name'] ]
	#df = df.convert_objects(convert_numeric=True)
	df[ ['progress', 'estimated-minutes'] ] = df[ ['progress', 'estimated-minutes'] ].apply(pd.to_numeric)
	df[ ['start-date', 'due-date'] ] = df[ ['start-date', 'due-date'] ].apply(pd.to_datetime)
	df['est-time'] = df['estimated-minutes']/60
	df.drop( 'estimated-minutes', axis=1, inplace=True ) 
	#is the task currently active? task.starttime > totime?, completed = False?
	df = df.loc[ df['start-date'] < totime, : ]
	df = df.loc[ df['completed'] == False, : ]
	#handle leave tasks; set progress and mark as completed
	#leave tasks are found in a project that has '15000' in its project-name 
	for i, row in df.iterrows():
		if ('150000' in row['project-name']) and ( 'leave' in row['content'] ):
			if row['due-date'].date() < totime:
				#mark as completed
				df.loc[i,'progress'] = 100
				r = requests.put( url + "tasks/" + str( row['id'] ) + "/complete.json", auth=HTTPBasicAuth( key, 'xxx' ), params=payload )
				print( 'mark as complete: ' + row['content'] + ' ' + str( r.ok ) )
			else:
				#set progress as ratio [b_days before totime*7.6 / estimated time in hours]
				Bdays_untilnow = len( pd.date_range(row['start-date'], totime, freq=pd.tseries.offsets.BDay()) )
				df.loc[i,'progress'] = np.round( 100 * 7.6 * Bdays_untilnow / row['est-time'] )
				payload = { 'todo-item': { 'progress': str(df.loc[i,'progress']) } }
				r = requests.put( url + "tasks/" + str( row['id'] ) + ".json", auth=HTTPBasicAuth( key, 'xxx' ), data=json.dumps(payload) )
				print( 'mark progress: ' + row['content'] + ' ' + str(df.loc[i,'progress']) + ' ' + str( r.ok ) )
				if not r.ok:
					print( r.raerar )
				#print( pd.date_range(row['start-date'], totime, freq=pd.tseries.offsets.BDay()) )
				#print( row )
				
	
	#assume task has start time, due date and estimated time, distribute time across period --->
	#pandas.tseries.offsets. 
	Bdays_in_target = pd.date_range(totime, nexttime, freq=pd.tseries.offsets.BDay())
	df.loc[ :,'est-time-next'] = 0
	for i, row in df.iterrows():
		if row['start-date'].year > 0 and row['due-date'].year > 0:
			if 1:																													#distribute remaining estimated time between today and end_date
				Bdays_in_period = pd.date_range(totime, row['due-date'], freq=pd.tseries.offsets.BDay())
			else:
				Bdays_in_period = pd.date_range(row['start-date'], row['due-date'], freq=pd.tseries.offsets.BDay())
			#all in past (task overdue) -> est_time_next   = row['est-time'] * (100 - row['progress'])/100
			#all in future (task not started) -> est_time_next   = row['est-time'] * (100 - row['progress'])/100
			#overlapping (task started before timenext and due after totime) -> modifier = overlapping_days / total_days
			est_time_next   = row['est-time'] * (100 - row['progress'])/100
			overlap = Bdays_in_period.intersection( Bdays_in_target )
			if overlap.size > 0:
				est_time_next = est_time_next * overlap.size / Bdays_in_period.size
			df.loc[i,'est-time-next'] = est_time_next
		else:
			df.loc[i,'est-time-next'] = row['est-time']
	#unassigned tasks have NaNs (and get grouped/displayed odd)
	df.loc[ df['responsible-party-lastname'].isnull(), 'responsible-party-lastname' ] = 'unassigned'
	#datetimes back to strings...
	df[ ['start-date', 'due-date'] ] = df[ ['start-date', 'due-date'] ].apply( lambda x: [ (y.strftime( "%b-%d-%Y") if pd.notnull(y) else '') for y in x] )
	#dftasks = df
	return df

#get all projects
def req_projects( url, companyid ):
	print( 'request for all projects within incas3' )
	payload = { "Content-Type": "application/json"}
	r = requests.get( url + "/companies/" + companyid + "/projects.json", auth=HTTPBasicAuth( key, 'xxx' ), params=payload )
	#print( r )
	print(  'page ' + r.headers['x-page'] + ' of ' + r.headers['x-pages'] )
	df = pd.DataFrame(r.json()['projects'])
	#df['short-name']= df['first-name'] + ' ' + df['last-name'].str[0] + '.'
	df = df[  ['name', 'id', 'startDate', 'endDate']  ]
	return df
	
#time totals; adds to dftasks and dfprojects	
def add_time_totals( url, totime, dfprojects, dftasks ):	
	print( 'request for time totals' )
	#add columns
	dfprojects.loc[ :,'total-hours-sum'] 			= 0
	dfprojects.loc[ :,'total-hours-estimated'] 		= 0
	dfprojects.loc[ :,'completed-hours-estimated'] 	= 0
	#request per project the total time logged
	for i, projectid in enumerate( dfprojects['id'] ):
		#hours before totime:
		print( dfprojects.loc[i, 'name'] )
		payload = { "Content-Type": "application/json" ,
					"toDate": datetime.date.strftime(totime, '%Y%m%d' ) }
		r = requests.get( url + "/projects/" + projectid + "/time/total.json", auth=HTTPBasicAuth( key, 'xxx' ), params=payload )
		dfprojects.loc[ i, 'total-hours-sum'] 			= r.json()['projects'][0]['time-totals']['total-hours-sum']
		dfprojects.loc[ i, 'total-hours-estimated']		= r.json()['projects'][0]['time-estimates']['total-hours-estimated']
		dfprojects.loc[ i, 'completed-hours-estimated']	= r.json()['projects'][0]['time-estimates']['completed-hours-estimated']
	
	#the same routine per task
	print( 'request for time totals' )
	dftasks.loc[ :,'total-hours-sum'] = 0;				thsloc = np.where( dftasks.columns=='total-hours-sum')[0]
	dftasks.loc[ :,'total-hours-estimated'] = 0;		theloc = np.where( dftasks.columns=='total-hours-estimated')[0]
	dftasks.loc[ :,'completed-hours-estimated'] = 0;	cheloc = np.where( dftasks.columns=='completed-hours-estimated')[0]
	for i, taskid in enumerate( dftasks['id'] ):
		print( dftasks.iloc[i]['project-name'][0:20] + '\t' + dftasks.iloc[i]['todo-list-name'][0:20] + '\t' + dftasks.iloc[i]['content'][0:10] )
		payload = { "Content-Type": "application/json" ,
					"toDate": datetime.date.strftime(totime, '%Y%m%d' ) }
		r = requests.get( url + "/tasks/" + str( taskid ) + "/time/total.json", auth=HTTPBasicAuth( key, 'xxx' ), params=payload )	
		dftasks.iloc[ i, thsloc]    = r.json()['projects'][0]['tasklist']['task']['time-totals']['total-hours-sum']
		dftasks.iloc[ i, theloc]	= r.json()['projects'][0]['tasklist']['task']['time-estimates']['total-hours-estimated']
		dftasks.iloc[ i, cheloc]	= r.json()['projects'][0]['tasklist']['task']['time-estimates']['completed-hours-estimated']	
	
#collecting data:
company_id 	= req_id( url )
dfpeople 	= req_people( url, company_id )
dftime		= req_logtime( url )
dftasks		= req_tasks( url, totime, nexttime )
dfprojects  = req_projects( url, company_id )
add_time_totals( url, totime, dfprojects, dftasks )	

#generating views
writer  = pd.ExcelWriter(outputfile, engine='xlsxwriter')		

if 1:		#Sheet 1:	give overview of logged time per person, pp-per project, pppp-per task
	table1   = pd.pivot_table( dftime, values = 'frac-time', index = ['person-last-name'], aggfunc=np.sum ,margins=False)
	table1   = pd.DataFrame( table1, index=table1.index )
	table1.to_excel(writer,'Sheet1', startrow=10, startcol=2)

	table2   = pd.pivot_table( dftime, values = 'frac-time', index = ['person-last-name', 'project-name'], aggfunc=np.sum ,margins=False)
	table2   = pd.DataFrame( table2, index=table2.index )
	table2.to_excel(writer,'Sheet1', startrow=10 + len( table1 ) + 4, startcol=1 )

	workbook  = writer.book
	worksheet = writer.sheets['Sheet1']
	worksheet.write( 0, 0, 'logged hours overview' )
	worksheet.write_row( 0, 1, ('from: ', datetime.date.strftime(fromtime, '%Y%m%d' ) ) )
	worksheet.write_row( 1, 1, ('to: ', datetime.date.strftime(totime, '%Y%m%d' ) ) )

	ljustfmt = workbook.add_format( {'align': 'right'} )
	datefmt  = workbook.add_format( {'num_format': 'd mmmm yyyy'} )
	numfmt   = workbook.add_format( {'num_format': '#.##'} )
	worksheet.set_column( 'A:C', 40, ljustfmt )
	worksheet.set_column( 'D:D', 40, numfmt )
	
	table3   = pd.pivot_table( dftime, values = 'frac-time', index = ['person-last-name', 'project-name', 'todo-item-name'], aggfunc=np.sum ,margins=False)
	table3   = pd.DataFrame( table3, index=table3.index )
	table3.to_excel(writer,'Sheet1', startrow= 10 + len( table1 ) + len( table2 ) + 8 )

if 1:		#Sheet 2:	give overview of estimated time per person, per project, per task
	
	table1   = pd.pivot_table( dftasks, values = 'est-time-next', index = ['responsible-party-lastname'], aggfunc=np.sum ,margins=False)
	table1   = pd.DataFrame( table1, index=table1.index )
	table1.to_excel(writer,'Sheet2', startrow=10, startcol=2)

	table2   = pd.pivot_table( dftasks, values = 'est-time-next', index = ['responsible-party-lastname', 'project-name'], aggfunc=np.sum ,margins=False)
	table2   = pd.DataFrame( table2, index=table2.index )
	table2.to_excel(writer,'Sheet2', startrow=10 + len( table1 ) + 4, startcol=1)
		
	table3 = dftasks[ ['responsible-party-lastname', 'project-name', 'content', 'est-time-next', 'est-time', 'progress', 'start-date', 'due-date','responsible-party-names'] ].set_index( ['responsible-party-lastname', 'project-name', 'content'] )
	table3.sortlevel(0, inplace=True)
	table3.to_excel(writer,'Sheet2', startrow=10 + len( table1 ) + len(table2) + 8, startcol=0)
	
	workbook  = writer.book
	worksheet = writer.sheets['Sheet2']
	worksheet.write( 0, 0, 'planned hours overview' )
	worksheet.write_row( 0, 1, ('from: ', datetime.date.strftime(totime, '%Y%m%d' ) ) )
	worksheet.write_row( 1, 1, ('to: ', datetime.date.strftime(nexttime, '%Y%m%d' ) ) )
	worksheet.write( 2, 0, 'progress taken into account when estimting time left on task' )
	worksheet.write( 3, 0, 'recurring tasks are not handled properly' )
	worksheet.write( 4, 0, 'time is not distributed over multiple persons' )
	worksheet.set_column( 'A:C', 40, ljustfmt )
	#worksheet.set_column( 'D:H', 20, ljustfmt )
	worksheet.set_column( 'G:H', 20, datefmt )
	worksheet.set_column( 'D:F', 20, numfmt )

if 1:		#Sheet 3:	overview per project	
	#table 1: projects -> hours logged until today, hours planned in total	
	table1 = dfprojects[ ['name', 'total-hours-sum','total-hours-estimated','completed-hours-estimated'] ].set_index( ['name'] )
	table1.sortlevel(0, inplace=True)
	table1.to_excel(writer,'Sheet3', startrow=10, startcol=2)

	#table 2: project;tasks -> total hours logged, hours planned, progress, end_date (ovedure)
	table2 = dftasks[ ['project-name', 'todo-list-name', 'content','total-hours-sum','total-hours-estimated','completed-hours-estimated'] ].set_index( ['project-name', 'todo-list-name', 'content'] )
	table2.sortlevel(0, inplace=True)
	table2.to_excel(writer,'Sheet3', startrow=10 + len( table1 ) + 8, startcol=0)
	
	workbook  = writer.book
	worksheet = writer.sheets['Sheet3']
	worksheet.write( 0, 0, 'projects time overview' )
	worksheet.set_column( 'A:E', 40, ljustfmt )

if 1:		#Sheet 4:	table with suggested timetell entries.
	grp = dftime[ ['person-last-name', 'project-name', 'date2', 'frac-time'] ].groupby( ['person-last-name', 'project-name', 'date2'] )
	table = grp.sum().unstack( level = -1 )	#put unique( date2 ) into columns, fileds are the sums of frac-time
	table = table.reset_index()						#remove the indexing
	table.insert( 1, 'project-number', table['project-name'].apply( lambda x: x[:6] ) )	#add column at position 2
	table.to_excel(writer,'Sheet4', startrow=0, startcol=0, merge_cells=False)
	
	#workbook  = writer.book
	worksheet = writer.sheets['Sheet4']
	#worksheet.write( 0, 0, 'Timetell pre-fill sheet' )
	worksheet.set_column( 'A:A', 5, numfmt )
	worksheet.set_column( 'B:C', 20, ljustfmt )
	worksheet.set_column( 'D:D', 40, ljustfmt )
	worksheet.set_column( 'E:P', 16, numfmt )
	
#write the excel sheet(s)
writer.save()	
#open the excel file
os.system( 'start excel.exe ' + outputfile )

