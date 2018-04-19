import xlrd
from xlrd import open_workbook, cellname
from xlrd.sheet import ctype_text   
import numpy as np
from numbers import Number
from decimal import Decimal

class processData:

	#For reading the excel file
	book = open_workbook("pollsSwedenUpdated.xls", "r")
	sheet_names = book.sheet_names()
	sheet = book.sheet_by_index(0)
	row1 = sheet.row(0)
	data = {}

	#for transforming the date, the format changes through the file
	sepdate = 41698
	months = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
	startYear = 2009
	startMonth = 2
	startDate = 28
	monthname = ["jan", "feb", "mars", "apr", "maj", "juni", "juli", "aug", "sep", "okt", "nov", "dec"]
	monthtomonth = {"spet" : "sep", "pub" : "ej publicerad" ,"jan" : "jan", "feb" : "feb", "mars" : "mars", "apr" : "apr", "maj" : "maj", "juni" : "juni", "juli" : "juli", "aug" : "aug", "sep" : "sep", "okt" : "okt", "nov" : "nov", "dec" : "dec", "sept" : "sep", "april": "apr", "Januari" : "jan", "Februari" : "feb", "Mars" : "mars", "April": "apr", "Maj" : "maj", "Juni" : "juni", "Juli" : "juli", "Augusti" : "aug", "September" : "sept", "Oktober" : "okt", "November" : "nov", "December" : "dec"}

	def readData(self):
		for i in range(0, 22):
			self.data[i] = []
		self.row1 = self.sheet.row(0)
		for i in range(2, 1632):
			self.row1 = self.sheet.row(i-1)
			for idx, cell_obj in enumerate(self.row1):
				cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
				#print('(%s) %s %s' % (idx, cell_type_str, cell_obj.value))
				if cell_obj.value == "" and idx == 0:
					#print("abc")
					break
				else:
					self.data[idx].append(cell_obj.value)


	def numberToDate(self, datenbr):
		diff = self.sepdate - datenbr
		



		self.sepdate -= diff
		if diff < 0:
			print("choklad")
			diff += 365
		elif diff > 2000:
			self.sepdate = 38976.0
			diff = 0
			self.startYear = 2006
			self.startMonth = 9
			self.startDate = 16
		#print("diff: " + str(diff), end="")
		self.startDate = self.startDate - diff
		# print("diff: " + str(diff) ,end=" ")
		while self.startDate <= 0:
			self.startMonth -= 1
			if self.startMonth == 2 and (self.startYear == 2004 or self.startYear == 2000 or self.startYear == 2008):
				self.startDate = self.months[self.startMonth-1] + 1 + self.startDate
			elif self.startMonth <= 0:
				self.startYear -= 1
				self.startMonth = 12
				self.startDate = self.months[self.startMonth-1] + self.startDate
			else:
				self.startDate = self.months[self.startMonth-1] + self.startDate

		# return (startDate, monthname[startMonth-1])

	def printDataDate(self):
		flag = True

		for i in range(0, len(self.data[16])):
			if isinstance(self.data[16][i], Number) and flag:
				flag = False
				self.sepdate = int(self.data[16][i])
			if str(self.data[16][i]) == "":
				print(str(self.data[0][i]) + " " + self.monthtomonth[str(self.data[1][i])])
			elif not isinstance(self.data[16][i], Number):
				str1 = self.data[16][i].split(" ")
				print(str(str1[0]) + " " + self.monthtomonth[str(str1[1]).lower()])
			else:
				self.numberToDate(int(self.data[16][i]))
				print(str(int(self.startDate)) + " " + str(self.monthname[self.startMonth-1])) 

	def distancedate(self, year, month):
		tempyear = 2018 - year
		tempmonth = 9 - self.monthname.index(month)

		return tempyear*12 + tempmonth


	def returnData(self):

		data = []

		#For now we ignore which company who did the study
		for i in range(0, len(self.data[16])):
			x.append([])
			x[i].append(int(self.date[0][i]))
			x[i].append(self.monthtomonth(str(self.date[1][i])))
			x[i].append(self.date[3][i])
			x[i].append(float(self.date[0][i]))



			#date
			flag = True
			for i in range(0, len(self.data[16])):
				if isinstance(self.data[16][i], Number) and flag:
					flag = False
					self.sepdate = int(self.data[16][i])
				if str(self.data[16][i]) == "":
					print(str(self.data[0][i]) + " " + self.monthtomonth[str(self.data[1][i])])
				elif not isinstance(self.data[16][i], Number):
					str1 = self.data[16][i].split(" ")
					print(str(str1[0]) + " " + self.monthtomonth[str(str1[1]).lower()])
				else:
					self.numberToDate(int(self.data[16][i]))
					print(str(int(self.startDate)) + " " + str(self.monthname[self.startMonth-1])) 


datafile = processData()

datafile.readData()
datafile.printDataDate()



# choices = {"jan": 1, "feb": 2, "mars": 3, "mar": 3, "april": 4, "apr": 4, }

# def monthToNumber(str):
# 	switch

# def fixdate(str):






	

# def monthToNumber(str):
# 	switch

# def fixdate(str):


# def monthToNumber(str):
# 	switch

# def fixdate(str):
