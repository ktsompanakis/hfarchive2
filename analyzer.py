# File name: analyzer.py
# Author: Konstantinos Tsompanakis
# Date created: 23/11/2015
# Date last modified: 20/01/2016
# Python Version: 2.7

# Begin code


import os
import xlwt
from xlrd import open_workbook
from xlwt import *
from xlutils.copy import copy
import ConfigParser
import time
import sys
import MMN_CuDo_Link
import time


debug = False

# Configure font style
style1 = xlwt.easyxf('font: bold 1')
style2 = xlwt.easyxf('font: color black;')
style3 = xlwt.easyxf('font: color blue;')


def deleteBlankLines(arrayOfMasks):	#Deletes blank lines 
	new_arr = []
	for line in arrayOfMasks:
		if not line.strip():
			continue
		else:
			new_arr.append(line)
	return new_arr


def deleteSpareLines(arrayOfMasks):	#Deletes spare lines
	new_arr = [] 
	for line in arrayOfMasks:
		if line.startswith('HEADER:') or line.startswith('DATA:'):
			continue
		else:
			new_arr.append(line)
	return new_arr


def lineSplitter(archive, arrayOfMasks):	#Splits masks inside an array
	if debug: print '\n~~~~~~~ lineSplitter func started ~~~~~~~~~\n'
	temp = []
	for line in archive:
		temp.append(line)
		if line.startswith('END'):
			arrayOfMasks.append(temp)
			temp = []
	return arrayOfMasks


def identifyMask(arrayOfMasks,captured_MB,captured_SN):	#Issolates simiral masks
	if debug: print '\n~~~~~~~ identifyMask func started ~~~~~~~~~\n'
	for mask in arrayOfMasks:
		try:
			if (mask[9][10:12] == 'MB') or (mask[7][10:12] == 'MB'):	# MB related masks
				captured_MB.append(mask)

			if (mask[9][10:12] == 'SN') or (mask[7][10:12] == 'SN'):  # SN related masks
				captured_SN.append(mask)

		except IndexError:
			pass


def snMaskAnalyze(captured_SN, book):
	if debug: print '\n~~~~~~~ snMaskAnalyze func started ~~~~~~~~~\n'
	print
	print "------------------------------------"
	print "------- analyzing SN mask ----------"
	print "------------------------------------"
	print
	sheet = book.add_sheet('SN mask',cell_overwrite_ok=True) #Create an xls sheet where we will store results

	snHeaders = ['Header', 'Date', 'Time', 'Message Group', 'Specific Mask', 'Type of Mask', 'MMN', 'Alarm Priority', 'Probable Cause', 'Specific Problem',
	 'Message Number', 'Mask Class', 'SN id', 'Unit', 'From', 'To', 'Supplementary Info 1', 'Supplementary Info 2', 'Supplementary Info 3',
	  'Supplementary Info 4']

	rows_included = 0

	for item in range(0,len(snHeaders)): #fill first raw of output excel with titles of Data
		sheet.write(0,item,snHeaders[item],style1)

	for mask in captured_SN:
		header = mask[0][0:38]
		date = mask[0][54:62]
		time = mask[0][64:72]
		messageGroup = mask[1][35:39]
		specificMask = mask[1][40:45]

		if mask[9][10:12] == 'SN':
			if mask[10][4:8] =='CONF':
				# EQUIPMENT ALARM SPECIFIC
				typeOfMask = mask[3][4:40]
				mmn = mask[3][62:67]
				alarmPriority = mask[4][22:35]
				probableCause = mask[5][22:45]
				specificProblem = mask[6][22:]
				messageNumber = mask[7][22:32]
				classOfMask = mask[9][10:12]
				SN_id = mask[9][21:23]
				transition_unit = mask[12][20:31]
				transition_from = mask[12][33:36]
				transition_to = mask[12][39:42]
				supplementaryInfo1 = mask[14][6:41]
				supplementaryInfo2 = mask[15][6:41]
				supplementaryInfo3 = mask[16][6:41]
				supplementaryInfo4 = mask[17][6:41]
			else:
				# OTHER TYPE OF EQUIPMENT ALARM
				typeOfMask = mask[3][4:40]
				mmn = mask[3][62:67]
				alarmPriority = mask[4][22:35]
				probableCause = mask[5][22:45]
				specificProblem = mask[6][22:]
				messageNumber = mask[7][22:32]
				classOfMask = mask[9][10:12]
				SN_id = mask[9][21:23]
				transition_unit = mask[9][10:15] + "-" + mask[9][21] + " -" + mask[9][36]
				transition_from = ""
				transition_to = ""
				supplementaryInfo1 = ""
				supplementaryInfo2 = ""
				supplementaryInfo3 = ""
				supplementaryInfo4 = ""

		#END OF EQUIPMENT ALARM SPECIFIC
		elif mask[7][10:12] == 'SN':
			typeOfMask = mask[2][4:40]
			probableCause = mask[3][22:45]
			specificProblem = mask[4][22:]
			messageNumber = mask[5][22:32]
			classOfMask = mask[7][10:12]
			SN_id = mask[7][21:23]
			transition_unit = mask[7][10:15] + "-" + mask[7][21] + " -" + mask[7][36]
			transition_from = ""
			transition_to = ""
			supplementaryInfo1 = ""
			supplementaryInfo2 = ""
			supplementaryInfo3 = ""
			supplementaryInfo4 = ""

		if debug:
			print 'header = ' + header
			print 'date = ' + date
			print 'time = ' + time
			print 'messageGroup = ' + messageGroup
			print 'specificMask = ' + specificMask
			print 'type of mask = ' + typeOfMask
			print 'mmn = ' + mmn
			print 'alarm priority = ' + alarmPriority
			print 'probable cause = ' + probableCause
			print 'specific problem = ' + specificProblem
			print 'message number = ' + messageNumber
			print 'class = ' + classOfMask
			print 'SN = ' + SN_id
			print 'unit = ' + transition_unit
			print 'from = ' + transition_from
			print 'to = ' + transition_to
			print 'supplementary info = ' + supplementaryInfo1
			print 'supplementary info = ' + supplementaryInfo2
			print 'supplementary info = ' + supplementaryInfo3
			print 'supplementary info = ' + supplementaryInfo4
			print '-----------------------------------------------'


		snInfo = [header, date, time, messageGroup, specificMask, typeOfMask, mmn, alarmPriority, probableCause, specificProblem, messageNumber,
			classOfMask, SN_id, transition_unit, transition_from, transition_to, supplementaryInfo1, supplementaryInfo2,
			supplementaryInfo3, supplementaryInfo4]

		for column in range(0, len(snInfo)):
			sheet.write(rows_included+1,column,snInfo[column],style2)


		mmnLink = MMN_CuDo_Link.mmnLinkUpdate(mmn)
		if mmnLink != '"NOTFOUND"':
			mmnStr = '"%s"'%mmn
			sheet.write(rows_included+1,6,xlwt.Formula('HYPERLINK(%s;%s)'%(mmnLink,mmnStr)),style3)

		rows_included+=1
	print str(rows_included) + ' SN entries exported'


def mbMaskAnalyze(captured_MB, book):
	if debug: print '\n~~~~~~~ mbMaskAnalyze func started ~~~~~~~~~\n'
	print
	print "------------------------------------"
	print "------- analyzing MB mask ----------"
	print "------------------------------------"
	print

	sheet = book.add_sheet('MB mask',cell_overwrite_ok=True) #Create an xls sheet where we will store results

	mbHeaders = ['Header', 'Date', 'Time', 'Message Group', 'Specific Mask', 'Type of Mask', 'MMN', 'Alarm Priority', 'Probable Cause', 'Specific Problem',
	 'Message Number', 'Mask Class', 'MB id', 'Unit', 'From', 'To', 'Supplementary Info 1', 'Supplementary Info 2', 'Supplementary Info 3',
	  'Supplementary Info 4']

	rows_included = 0

	for item in range(0,len(mbHeaders)): #fill first raw of output excel with titles of Data
		sheet.write(0,item,mbHeaders[item],style1)

	for mask in captured_MB:
		header = mask[0][0:38]
		date = mask[0][54:62]
		time = mask[0][64:72]
		messageGroup = mask[1][35:39]
		specificMask = mask[1][40:45]

		if mask[9][10:12] == 'MB':
			if mask[10][4:8] =='CONF':
				# EQUIPMENT ALARM SPECIFIC
				typeOfMask = mask[3][4:40]
				mmn = mask[3][62:67]
				alarmPriority = mask[4][22:35]
				probableCause = mask[5][22:45]
				specificProblem = mask[6][22:]
				messageNumber = mask[7][22:32]
				classOfMask = mask[9][10:12]
				MB_id = mask[9][21:23]
				transition_unit = mask[12][20:31]
				transition_from = mask[12][33:36]
				transition_to = mask[12][39:42]
				supplementaryInfo1 = mask[14][6:41]
				supplementaryInfo2 = mask[15][6:41]
				supplementaryInfo3 = mask[16][6:41]
				supplementaryInfo4 = mask[17][6:41]
			else:
				# OTHER TYPE OF EQUIPMENT ALARM
				typeOfMask = mask[3][4:40]
				mmn = mask[3][62:67]
				alarmPriority = mask[4][22:35]
				probableCause = mask[5][22:45]
				specificProblem = mask[6][22:]
				messageNumber = mask[7][22:32]
				classOfMask = mask[9][10:12]
				MB_id = mask[9][21:23]
				if mask[9][10:14] == "MB  ":	# MBIC only
					transition_unit = "MBIC" + " -" + mask[9][21]
				else:
					transition_unit = mask[9][10:14] + " -" + mask[9][21] + " -" + mask[9][36]
				transition_from = ""
				transition_to = ""
				supplementaryInfo1 = ""
				supplementaryInfo2 = ""
				supplementaryInfo3 = ""
				supplementaryInfo4 = ""

		#END OF EQUIPMENT ALARM SPECIFIC
		elif mask[7][10:12] == 'MB':
			typeOfMask = mask[2][4:40]
			probableCause = mask[3][22:45]
			specificProblem = mask[4][22:]
			messageNumber = mask[5][22:32]
			classOfMask = mask[7][10:12]
			MB_id = mask[7][21:23]
			if mask[7][10:14] == "MB  ":	# MBIC only
				transition_unit = "MBIC" + " -" + mask[7][21]
			else:
				transition_unit = mask[7][10:14] + " -" + mask[7][21] + " -" + mask[7][36]
			transition_from = ""
			transition_to = ""
			supplementaryInfo1 = ""
			supplementaryInfo2 = ""
			supplementaryInfo3 = ""
			supplementaryInfo4 = ""


		if debug:
			print 'header = ' + header
			print 'date = ' + date
			print 'time = ' + time
			print 'messageGroup = ' + messageGroup
			print 'specificMask = ' + specificMask
			print 'type of mask = ' + typeOfMask
			print 'mmn = ' + mmn
			print 'alarm priority = ' + alarmPriority
			print 'probable cause = ' + probableCause
			print 'specific problem = ' + specificProblem
			print 'message number = ' + messageNumber
			print 'class = ' + classOfMask
			print 'MB = ' + MB_id
			print 'unit = ' + transition_unit
			print 'from = ' + transition_from
			print 'to = ' + transition_to
			print 'supplementary info = ' + supplementaryInfo1
			print 'supplementary info = ' + supplementaryInfo2
			print 'supplementary info = ' + supplementaryInfo3
			print 'supplementary info = ' + supplementaryInfo4
			print '-----------------------------------------------'


		mbInfo = [header, date, time, messageGroup, specificMask, typeOfMask, mmn, alarmPriority, probableCause, specificProblem, messageNumber,
			classOfMask, MB_id, transition_unit, transition_from, transition_to, supplementaryInfo1, supplementaryInfo2,
			supplementaryInfo3, supplementaryInfo4]

		for column in range(0, len(mbInfo)):
			sheet.write(rows_included+1,column,mbInfo[column],style2)


		mmnLink = MMN_CuDo_Link.mmnLinkUpdate(mmn)
		if mmnLink != '"NOTFOUND"':
			mmnStr = '"%s"'%mmn
			sheet.write(rows_included+1,6,xlwt.Formula('HYPERLINK(%s;%s)'%(mmnLink,mmnStr)),style3)

		rows_included+=1
	print str(rows_included) + ' MB entries exported'


def main():
	book = Workbook()
	os.system('cls')
	while True:
		print " /-----------------------------\\"
		print "|~~~~~ HF.ARCHIVE ANALYZER ~~~~~|"
		print "|~~~~~     version 0.5     ~~~~~|"
		print " \-----------------------------/"
		print
		print "Please choose file name bellow: (only UTF-8 encoding is supported)"
		print
		print "1. HF.ARCHIVE.txt, (" + os.getcwd() + ")"
		print "2. Choose filename manualy"
		print "3. Quit"
		print
		userChoice = raw_input('Enter your choice: ')
		if userChoice == '1':
			filename = "HF.ARCHIVE.txt"
			print
			break
		elif userChoice == '2':
			filename = raw_input('Enter filename: ')
			break
		elif userChoice == '3':
			print 'Exiting... '
			print
			time.sleep(1)
			sys.exit()
			break
		else:
			print 'Sorry you have enter an invalid choice... Please try again: '

	# Read from Config file
	try:
		config = ConfigParser.ConfigParser()
		config.read('config.ini')
		global debug
		debug = config.getboolean('Parameters', 'debug')	#Debug mode
	except ConfigParser.ParsingError, err:
		# print 'Could not parse:', err
		# sys.exit()
		pass
	except ConfigParser.NoSectionError, err:
		# print 'Could not parse:', err
		# sys.exit()
		pass



	try:
		f = open("%s" %filename, "r")
		pass
	except IOError:
		print "File does not exist."
		time.sleep(1)
		print "Quiting..."
		time.sleep(1)
		sys.exit()

	arrayOfMasks = []
	captured_MB = []
	captured_SN = []

	archive1 = f.readlines()
	f.close()

	#mark as comment any line bellow you don't want to execute
	archive2 = deleteBlankLines(archive1)		# Delete blank lines from input file
	archive = deleteSpareLines(archive2)		# Delete lines with HEADER or DATA text
	lineSplitter(archive, arrayOfMasks)			# Split masks into list
	identifyMask(arrayOfMasks,captured_MB,captured_SN)	# Identify mask type
	mbMaskAnalyze(captured_MB, book)			# Analyze MB masks
	snMaskAnalyze(captured_SN, book)			# Analyze SN masks


	current_time = time.strftime("%d.%m.%y_%H.%M", time.localtime())
	output_name = "ANALYZED_"+ current_time + ".xls"

	print
	try:
		book.save(output_name) #Export output in a new xls file
	except IOError:
		print
		print "ERROR"
		print "Output file is busy."
		print "Please close " + output_name + " file and try again"
		time.sleep(10)
		sys.exit()
	else:
		print 'Output is saved to ' + output_name

	if debug: print '\n~~~~~~~~~~ program end ~~~~~~~~~~~\n'
	u = raw_input('Press Enter to exit... ')

if __name__ == '__main__':
	main()
