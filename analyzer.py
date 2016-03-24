import os
import xlwt
from xlrd import open_workbook
from xlwt import *
from xlutils.copy import copy
import ConfigParser
import time
import sys
import MMN_CuDo_Link


debug = False

# Configure font style
style1 = xlwt.easyxf('font: bold 1')
style2 = xlwt.easyxf('font: color black;')
style3 = xlwt.easyxf('font: color blue;')


def lineSplitter(f, t, temp):	#Splits masks inside an array
	if debug: print '\n~~~~~~~ lineSplitter func started ~~~~~~~~~\n'
	for line in f:
		# if line.startswith('\r') or line.startswith('\xef') or line.startswith('\n'):
		# 	continue
		temp.append(line)
		if line.startswith('END'):
			t.append(temp)
			temp = []
	f.close()

	t = deleteBlankLines(t)
	return t


def deleteBlankLines(t):	#Deletes blank lines in the begin of each mask
	for mask in t:
		i = 0
		for line in mask:
			if line == "HEADER:":
				break
			if not line.strip():
				del mask[i]
			i+=1
	return t


def issolateHeader(t,captured):	#Issolates simiral masks
	if debug: print '\n~~~~~~~ issolateHeader func started ~~~~~~~~~\n'
	for mask in t:
		try:
			#print mask[12][10:12]
			if (mask[12][10:12] == 'MB') or (mask[10][10:12] == 'MB'):	# MB related masks
				captured.append(mask)
				# for line in mask:
				# 	if line[10:12] == 'MB':
				# 		captured.append(mask)

		except IndexError:
			# print "exception captured"
			pass


def mbMaskAnalyze(captured, book):
	#print captured[1][1]
	if debug: print '\n~~~~~~~ mbMaskAnalyze func started ~~~~~~~~~\n'
	print
	print "------------------------------------"
	print "------- analyzing MB mask ----------"
	print "------------------------------------"
	print 
	# for mask in range(captured):
	# 	for line in range(mask):
	# 		print captured[mask][line]

	# for mask in range(0,len(captured)):
	# 	for line in mask:
	# 		print captured[mask][line]

	sheet = book.add_sheet('MB mask',cell_overwrite_ok=True) #Create an xls sheet where we will store results

	mbHeaders = ['Header', 'Date', 'Time', 'Message Group', 'Specific Mask', 'Type of Mask', 'MMN', 'Alarm Priority', 'Probable Cause', 'Specific Problem',
	 'Message Number', 'Mask Class', 'MB id', 'Unit', 'From', 'To', 'Supplementary Info 1', 'Supplementary Info 2', 'Supplementary Info 3',
	  'Supplementary Info 4']
	
	rows_included = 0

	for item in range(0,len(mbHeaders)): #fill first raw of output excel with titles of Data
		sheet.write(0,item,mbHeaders[item],style1)

	for mask in captured:
		header = mask[2][0:31]
		date = mask[2][54:62]
		time = mask[2][64:72]
		messageGroup = mask[3][35:39]
		specificMask = mask[3][40:45]

		if mask[12][10:12] == 'MB':
			if mask[13][4:8] =='CONF':
				# EQUIPMENT ALARM SPECIFIC			
				typeOfMask = mask[6][4:40]
				mmn = mask[6][62:67]
				alarmPriority = mask[7][22:35]
				probableCause = mask[8][22:45]
				specificProblem = mask[9][22:]
				messageNumber = mask[10][22:32]
				classOfMask = mask[12][10:12]
				MB_id = mask[12][21:23]
				transition_unit = mask[15][20:31]
				transition_from = mask[15][33:36]
				transition_to = mask[15][39:42]
				supplementaryInfo1 = mask[17][6:41]
				supplementaryInfo2 = mask[18][6:41]
				supplementaryInfo3 = mask[19][6:41]
				supplementaryInfo4 = mask[20][6:41]
			else:
				# OTHER TYPE OF EQUIPMENT ALARM
				typeOfMask = mask[6][4:40]
				mmn = mask[6][62:67]
				alarmPriority = mask[7][22:35]
				probableCause = mask[8][22:45]
				specificProblem = mask[9][22:]
				messageNumber = mask[10][22:32]
				classOfMask = mask[12][10:12]
				MB_id = mask[12][21:23]
				if mask[12][10:14] == "MB  ":	# MBIC only
					transition_unit = "MBIC" + " -" + mask[12][21]
				else:
					transition_unit = mask[12][10:14] + " -" + mask[12][21] + " -" + mask[12][36]
				transition_from = ""
				transition_to = ""
				supplementaryInfo1 = ""
				supplementaryInfo2 = ""
				supplementaryInfo3 = ""
				supplementaryInfo4 = ""

		#END OF EQUIPMENT ALARM SPECIFIC
		elif mask[10][10:12] == 'MB':
			typeOfMask = mask[5][4:40]
			probableCause = mask[6][22:45]
			specificProblem = mask[7][22:]
			messageNumber = mask[8][22:32]
			classOfMask = mask[10][10:12]
			MB_id = mask[10][21:23]
			if mask[10][10:14] == "MB  ":	# MBIC only
				transition_unit = "MBIC" + " -" + mask[10][21]
			else:
				transition_unit = mask[10][10:14] + " -" + mask[10][21] + " -" + mask[10][36]
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
	print str(rows_included) + ' entries exported'


def linePrinter(t):		#Prints all masks with new entry indicator between them
	for row in t:
		for column in row:
			print column
		print '-----------------NEW ENTRY----------------------------'


def fileWriter(t):		#Writes all masks to a file
	a = open("log.txt", "a")
	for row in t:
		for column in row:
			a.write(column)
		a.write('\n-------------------NEW ENTRY-------------------------\n')
	a.close()


def main():
	book = Workbook()
	while True:
		print " /-----------------------------\\"
		print "|~~~~~ HF.ARCHIVE ANALYZER ~~~~~|"
		print "|~~~~~     version 0.3     ~~~~~|"
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

	# f = open("%s" %filename, "r")
	t = []
	temp = []
	captured = []

	#mark as comment any line bellow you don't want to execute
	lineSplitter(f, t, temp)
	issolateHeader(t,captured)
	#mbPrinter(captured)
	# linePrinter(t)
	mbMaskAnalyze(captured, book)
	#fileWriter(t)

	#check(t)
	
	try:
		book.save('ANALYZED.xls') #Export output in a new xls file
	except IOError:
		print
		print "ERROR"
		print "Output file is busy."
		print "Please close ANALYZED.xls file and try again"
		time.sleep(10)
		sys.exit()
	else:
		print 'Output is saved to ANALYZED.xls'

	if debug: print '\n~~~~~~~~~~ program end ~~~~~~~~~~~\n'
	u = raw_input('Press Enter to exit... ')

if __name__ == '__main__':
	main()
