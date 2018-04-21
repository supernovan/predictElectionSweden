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
	monthtomonth = {"spet" : "sep", "pub" : "ej publicerad" ,"jan" : "jan", "feb" : "feb", "mars" : "mars", "apr" : "apr", "maj" : "maj", "juni" : "juni", "juli" : "juli", "aug" : "aug", "sep" : "sep", "okt" : "okt", "nov" : "nov", "dec" : "dec", "sept" : "sep", "april": "apr", "Januari" : "jan", "Februari" : "feb", "Mars" : "mars", "April": "apr", "Maj" : "maj", "Juni" : "juni", "Juli" : "juli", "Augusti" : "aug", "September" : "sep", "Oktober" : "okt", "November" : "nov", "December" : "dec"}
	electiondates = [(2002, 15), (2006, 17), (2010, 19), (2014, 14)]
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
			#print("choklad")
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
		#print(str(month))
		tempyear = 2018 - year
		newmonth = self.monthtomonth[month]
		tempmonth = 8 - self.monthname.index(newmonth)

		return tempyear*12 + tempmonth

	def checkData(self, element):
		#print(element)
		if not element:
			return 0.0
		else:
			tempelement = str(element).replace(",", ".")
			#print(tempelement)
			return float(tempelement)

	def datetoelection(self, date):
		if len(date) == 2:
			idx = 4
			#print(" " + str(self.monthname.index((date[1]))), end="")
			for i in range(0, len(self.electiondates)):
				if self.electiondates[i][0] - float(date[0]) > 0:
					idx = i
					break
				elif self.electiondates[i][0] - float(date[0]) == 0.0 and self.monthname.index((date[1])) - 8 <= 0:
					idx = i
					break

			return idx
			#print(" " + str(idx))
		elif len(date) == 3:
			idx = 4
			#print(" " + str(self.monthname.index((date[1]))), end="")
			for i in range(0, len(self.electiondates)):
				if self.electiondates[i][0] - float(date[0]) > 0:
					idx = i
					break
				elif self.electiondates[i][0] - float(date[0]) == 0.0 and self.monthname.index((date[1])) - 8 < 0:
					idx = i
					break
				elif self.electiondates[i][0] - float(date[0]) == 0.0 and self.monthname.index((date[1])) - 8 == 0 and  int(date[2]) - self.electiondates[i][1] < 0.2:
					idx = i
					break
			#print(" " + str(idx))
			return idx
	def returnData(self):

		tempdata = []

		flag = True
		#For now we ignore which company who did the study
		for i in range(0, len(self.data[16])):
			tempdata.append([])
			#print(str(self.data[0][i]) + " " +  str(self.data[1][i]))
			tempdata[i].append(self.distancedate(int(self.data[0][i]), str(self.data[1][i])))
			tempdata[i].append(self.checkData(self.data[3][i]))
			tempdata[i].append(self.checkData(self.data[4][i]))
			tempdata[i].append(self.checkData(self.data[5][i]))
			tempdata[i].append(self.checkData(self.data[6][i]))
			tempdata[i].append(self.checkData(self.data[7][i]))
			tempdata[i].append(self.checkData(self.data[8][i]))
			tempdata[i].append(self.checkData(self.data[9][i]))
			tempdata[i].append(self.checkData(self.data[10][i]))
			tempdata[i].append(self.checkData(self.data[11][i]))
			tempdata[i].append(self.checkData(self.data[15][i]))
			nbrplebs = self.checkData(self.data[18][i])
			if not nbrplebs:
				nbrplebs = 1.0
			else:
				nbrplebs / 1000.0
			if self.checkData(self.data[21][i]):
				nbrplebs = self.checkData(self.data[21][i])/1000.0
			tempdata[i].append(nbrplebs)

			#get which election result it belongs to
			
			date = []
			if isinstance(self.data[16][i], Number) and flag:
				flag = False
				self.sepdate = int(self.data[16][i])
			if str(self.data[16][i]) == "":
				#print(str(self.data[0][i]) + " " + self.monthtomonth[str(self.data[1][i])], end="")
				#print(self.monthtomonth[str(self.data[1][i])])
				date.append(str(self.data[0][i]))
				date.append(self.monthtomonth[str(self.data[1][i])])
			elif not isinstance(self.data[16][i], Number):
				str1 = self.data[16][i].split(" ")
				#print(str(self.data[0][i]) + " " + str(str1[0]) + " " + self.monthtomonth[str(str1[1]).lower()], end="")
				date.append(str(self.data[0][i]))
				date.append(self.monthtomonth[str(str1[1]).lower()])
				date.append(str(str1[0]))
			else:

				self.numberToDate(int(self.data[16][i]))
				#print(str(self.data[0][i]) + " " + str(int(self.startDate)) + " " + str(self.monthname[self.startMonth-1]), end="") 
				date.append(str(self.data[0][i]))
				date.append(str(self.monthname[self.startMonth-1]))
				date.append(str(int(self.startDate)))
			
			tempdata[i].append(self.datetoelection(date))

		return tempdata

datafile = processData()

datafile.readData()
x = datafile.returnData()
#print(x)

result = {}
result[0] = [15.2, 26.2, 30.1, 23.3] #moderaterna
result[1] = [13.4, 7.5, 7.1, 5.4] #liberaler
result[2] = [6.2, 7.9, 6.6, 6.1] #centern
result[3] = [9.1, 6.6, 5.6, 4.6] #kd
result[4] = [39.8, 35.0, 30.7, 31.0] #sossarna
result[5] = [8.4, 5.9, 5.6, 5.7] #vänster
result[6] = [4.6, 5.2, 7.3, 6.9] #miljöpartiet
result[7] = [1.4, 2.9, 5.7, 12.9] #Sverigedemokraterna
result[8] = [1.7, 2.8, 1.4, 4.0] #övriga


yfuture = []
xfuture = []
xtraining = []
ytraining = []
ytemp = []
for i in range(0, 4):
	ytemp.append([])
	for j in range(0, 9):
		ytemp[i].append(result[j][i])

for row in x:
	if row[len(row)-1] == 4:
		xfuture.append(row)
	elif row[len(row)-1] == 3:
		xtraining.append(row[:-1])
		ytraining.append(ytemp[3])
	elif row[len(row)-1] == 2:
		xtraining.append(row[:-1])
		ytraining.append(ytemp[2])
	elif row[len(row)-1] == 1:
		xtraining.append(row[:-1])
		ytraining.append(ytemp[1])
	elif row[len(row)-1] == 2:
		xtraining.append(row[:-1])
		ytraining.append(ytemp[0])

#print(ytraining)


#fit model

import numpy as np
from keras.models import Sequential
from keras.layers import Dense
from sklearn.preprocessing import normalize
from keras.layers import Embedding
from keras.layers import LSTM
from keras.layers import Dropout
from keras.layers import Activation

xforkeras = np.array(xtraining)
yforkeras = np.array(ytraining)
print(len(x[0]))

# model = Sequential()
# model.add(Dense(80, input_dim=12, activation='tanh'))
# model.add(Dense(80, activation='tanh'))
# model.add(Dense(9, activation='linear'))
# model.compile(loss='mse', optimizer='adam', metrics=['accuracy'])

# model.fit(xforkeras, yforkeras, epochs=650, batch_size=50)
# scores = model.evaluate(xforkeras, yforkeras)
# print("\n%s: %.2f%%" % (model.metrics_names[1], scores[1]*100))
dropout = 0.25
neurons = 200
print(str(xforkeras.shape[0]) + " " + str(xforkeras.shape[1]))
model = Sequential()
model.add(LSTM(neurons, return_sequences=True, input_shape=(12, 902), activation="tanh"))
model.add(Dropout(dropout))
model.add(LSTM(neurons, return_sequences=True, activation="tanh"))
model.add(Dropout(dropout))
model.add(LSTM(neurons, activation="tanh"))
model.add(Dropout(dropout))
model.add(Dense(units=9))
model.add(Activation("linear"))
model.compile(loss='mse', optimizer='adam', metrics=['accuracy'])

model.fit(xforkeras, yforkeras, epochs=650, batch_size=50)
scores = model.evaluate(xforkeras, yforkeras)
print("\n%s: %.2f%%" % (model.metrics_names[1], scores[1]*100))




#try to predict election 2018
xtest = np.array((xfuture[0][:-1], xfuture[1][:-1]))
# print(x_lastpoll)


yhat = model.predict(xtest, verbose=0)


print("M : " + str(yhat[0][0]))
print("L : " + str(yhat[0][1]))
print("C : " + str(yhat[0][2]))
print("KD : " +str( yhat[0][3]))
print("S : " + str(yhat[0][4]))
print("V : " + str(yhat[0][5]))
print("Mp : " +str( yhat[0][6]))
print("SD : " +str( yhat[0][7]))
print("Ö : " + str(yhat[0][8]))

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
