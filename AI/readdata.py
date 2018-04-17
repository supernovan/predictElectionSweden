import xlrd
from xlrd import open_workbook, cellname
from xlrd.sheet import ctype_text   
import numpy as np

book = open_workbook("pollsSwedenUpdated.xls", "r")
sheet_names = book.sheet_names()


sheet = book.sheet_by_index(0)

row1 = sheet.row(0)

data = {}

for i in range(0, 22):
	data[i] = []
row1 = sheet.row(0)
for i in range(2, 1632):
	row1 = sheet.row(i-1)
	for idx, cell_obj in enumerate(row1):
		cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
		#print('(%s) %s %s' % (idx, cell_type_str, cell_obj.value))
		if cell_obj.value == "" and idx == 0:
			#print("abc")
			break
		else:
			data[idx].append(cell_obj.value)

# choices = {"jan": 1, "feb": 2, "mars": 3, "mar": 3, "april": 4, "apr": 4, }

# def monthToNumber(str):
# 	switch

# def fixdate(str):
from numbers import Number
from decimal import Decimal
sepdate = 41698
months = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
startYear = 2009
startMonth = 2
startDate = 28

monthname = ["jan", "feb", "mars", "april", "maj", "juni", "juli", "aug", "sept", "okt", "nov", "dec"]
monthtomonth = {"Januari" : "jan", "Februari" : "feb", "Mars" : "mars", "April": "april", "Maj" : "maj", "Juni" : "juni", "Juli" : "juli", "Augusti" : "aug", "September" : "sept", "Oktober" : "okt", "November" : "nov", "December" : "dec"}
def numberToDate(datenbr):
	global sepdate, months, startYear, startMonth, startDate, monthname
	diff = sepdate - datenbr
	



	sepdate -= diff
	if diff < 0:
		print("choklad")
		diff += 365
	elif diff > 2000:
		sepdate = 38976.0
		diff = 0
		startYear = 2006
		startMonth = 9
		startDate = 16
	#print("diff: " + str(diff), end="")
	startDate = startDate - diff
	print("diff: " + str(diff) ,end=" ")
	while startDate <= 0:
		startMonth -= 1
		if startMonth == 2 and (startYear == 2004 or startYear == 2000 or startYear == 2008):
			startDate = months[startMonth-1] + 1 + startDate
		elif startMonth <= 0:
			startYear -= 1
			startMonth = 12
			startDate = months[startMonth-1] + startDate
		else:
			startDate = months[startMonth-1] + startDate

	return (startDate, monthname[startMonth-1])

flag = True

for i in range(0, len(data[16])):
	if isinstance(data[16][i], Number) and flag:
		flag = False
		sepdate = int(data[16][i])
	if str(data[16][i]) == "":
		print(str(data[0][i]) + " " + monthtomonth[str(data[1][i])])
	elif not isinstance(data[16][i], Number):
		print(data[16][i])
	else:
		test = numberToDate(int(data[16][i]))
		print(str(test[0]) + " " + test[1]  + " " + str(startYear) + " " + str(data[16][i])) 
	

# def monthToNumber(str):
# 	switch

# def fixdate(str):


# def monthToNumber(str):
# 	switch

# def fixdate(str):
