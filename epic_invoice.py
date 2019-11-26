#import xls files
import requests
import openpyxl


def GetInputFileNames(path):
	from os import walk
	shanksNAGFiles = []
	for (dirpath, dirnames, filenames) in walk(path):
		for filename in filenames:
			if not filename.startswith('~'): 
				shanksNAGFiles.append(filename)
	return shanksNAGFiles

def CalcEpicsHours(sheet):
	nameCellX = 'F'
	hoursCellX = 'C'
	epicHours = {}
	# starting from 2 since 1 is the header.
	for x in range(2, 10000):
		nameCell = sheet[nameCellX + str(x)]
		hoursCell = sheet[hoursCellX + str(x)]
		if nameCell.value:
			if not nameCell.value in epicHours:
				epicHours[nameCell.value] = 0
			epicHours[nameCell.value] = epicHours[nameCell.value] + float(hoursCell.value)
	return epicHours

def InsertHours(sheet, hours):
	nameCellX = 'B'
	hoursCellX = 'E'
	for x in range(6, 34):
		nameCell = sheet[nameCellX + str(x)]
		hoursCell = sheet[hoursCellX + str(x)]
		if nameCell.value:
			if not nameCell.value in epicHours:
				sheet[hoursCellX + str(x)] = 0
				print ('no hours for:' + nameCell.value)
			else:
				sheet[hoursCellX + str(x)] = epicHours[nameCell.value]
				print (str(epicHours[nameCell.value]) + ' hours for:' + nameCell.value)


outputFileNameTemplate = 'epic invoice calc template.xlsx'

shanksNAGPath = 'C:\\Python\\Projects\\Python2\\Shanks2\\NAG Invoice 2'
shanksNLSPath = 'C:\\Python\\Projects\\Python2\\Shanks2\\NLS Invoice 2'
steakholdersPath = 'C:\\Python\\Projects\\Python2\\Steakholders2'
itmsPath = ''

#get input file names
#path = shanksNAGPath
#path = shanksNLSPath
path = steakholdersPath
files = GetInputFileNames(path)
for fileName in files:

	#get sheet
	inputFileName = path + '\\' + fileName
	print('++++++++++++++++++++++')
	print(inputFileName)
	print('++++++++++++++++++++++')
	book = {}
	book[inputFileName] = openpyxl.load_workbook(inputFileName)
	sheet = book[inputFileName]['Worklogs']

	#get hours
	epicHours = CalcEpicsHours(sheet)

	for employee in epicHours:
		print(employee + ':' + str(epicHours[employee]))

	print('_______')

	exportBook = openpyxl.load_workbook(outputFileNameTemplate)

	exportSheet = exportBook['Renumeration']

	InsertHours(exportSheet, epicHours)

	outputFileName = path + '\\' + fileName[0:(len(fileName) - 5)] + ' invoice calc.xlsx'
	exportBook.save(outputFileName)

