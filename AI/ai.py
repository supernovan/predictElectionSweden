import xlrd
from xlrd import open_workbook, cellname
from xlrd.sheet import ctype_text   
import numpy as np

book = open_workbook("pollsSweden.xls", "r")
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


# for i in range(0, 1):
# 	print(data[0])







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




#INPUT
# Electiondate - Polldate. If the poll is after the election date 
# it's trying to fit into the next election date result.

# Which company who made the poll, I use 1-C coding for the companies

# what each party, those who doesn't know and other smaller parties.

# from keras.models import Sequential
# from keras.layers import Dense

#calculating number of companies for the input 
companies = {}
for x in data[2]:
	if x not in companies:
		companies[x] = len(companies)
# print(len(companies))
# print(companies)
xven = []
for i in range(0, len(data[3])):
	xven.append([])
	xven[i].append(0) #date
	if str(data[3][i]) != "":
		xven[i].append(float(str(data[3][i]).replace(",", ".")))
	else:
		xven[i].append(0.0)
	if str(data[4][i]) != "":
		xven[i].append(float(str(data[4][i]).replace(",", ".")))
	else:
		xven[i].append(0.0)
	if str(data[5][i]) != "":
		xven[i].append(float(str(data[5][i]).replace(",", ".")))
	else:
		xven[i].append(0.0)
	if str(data[6][i]) != "":
		xven[i].append(float(str(data[6][i]).replace(",", ".")))
	else:
		xven[i].append(0.0)
	if str(data[7][i]) != "":
		xven[i].append(float(str(data[7][i]).replace(",", ".")))
	else:
		xven[i].append(0.0)
	if str(data[8][i]) != "":
		xven[i].append(float(str(data[8][i]).replace(",", ".")))
	else:
		xven[i].append(0.0)
	if str(data[9][i]) != "":
		xven[i].append(float(str(data[9][i]).replace(",", ".")))
	else:
		xven[i].append(0.0)

	# xven[i].append(float(str(data[3][i]).replace(",", ".")))
	# xven[i].append(float(str(data[4][i]).replace(",", ".")))
	# xven[i].append(float(str(data[5][i]).replace(",", ".")))
	# xven[i].append(float(str(data[6][i]).replace(",", ".")))
	# xven[i].append(float(str(data[7][i]).replace(",", ".")))
	# xven[i].append(float(str(data[8][i]).replace(",", ".")))
	# xven[i].append(float(str(data[9][i]).replace(",", ".")))
	if str(data[10][i]) != "":
		xven[i].append(float(str(data[10][i]).replace(",", ".")))
	else:
		xven[i].append(0.0)
	if str(data[11][i]) != "":
		xven[i].append(float(str(data[11][i]).replace(",", ".")))
	else:
		xven[i].append(0.0)
	
	if (data[15][i] != ""):
		xven[i].append(float(data[15][i]))
	else:
		xven[i].append(0.0)
	index = len(xven)
	for j in range(0, 32):
		#print(data[2][i] + " : " + str(companies.get(data[2][i])))
		if j == companies[data[2][i]]:

			xven[i].append(1)

		else:
			xven[i].append(0)





electiondates = [14, 19, 17, 15]

trainvec = []
for i in range(0, 9):
	trainvec.append([])

#print(len(trainvec))
#print(data[16])
idx = 0
for (x, y) in zip(data[0], data[16]):
	flag = -1
	before = len(str(y))
	newy = str(y).replace("-", " ")
	

	timefrom = 0.0
	#print(newy)
	date = newy.split(" ")
	if before != len(str(newy)):
		print(str(date))
		date.pop(0)
		print(date)
	for i in range (0, 4):
		# print(x)
		# print(date)
		if float(x) - (2014-i*4) > 0:
			timefrom = float(x) - (2014-i*4)
			flag = i
			break
		elif float(x) - (2014-i*4) == 0:
			if len(date) == 1:
				timefrom = 0.5
				flag = i
				break
			elif float(date[0]) - electiondates[i] < 0 and not (date[1] == "sep" or date[1] == "okt" or date[1] == "nov" or date[1] == "dec"):
				timefrom = 0.2
				flag = i
				break
	if flag != -1:
		trainvec[0].append(result[0][3-flag])
		trainvec[1].append(result[1][3-flag])
		trainvec[2].append(result[2][3-flag])
		trainvec[3].append(result[3][3-flag])
		trainvec[4].append(result[4][3-flag])
		trainvec[5].append(result[5][3-flag])
		trainvec[6].append(result[6][3-flag])
		trainvec[7].append(result[7][3-flag])
		trainvec[8].append(result[8][3-flag])
	else:
		if int(x) < 2002:
			trainvec[0].append(result[0][0])
			trainvec[1].append(result[1][0])
			trainvec[2].append(result[2][0])
			trainvec[3].append(result[3][0])
			trainvec[4].append(result[4][0])
			trainvec[5].append(result[5][0])
			trainvec[6].append(result[6][0])
			trainvec[7].append(result[7][0])
			trainvec[8].append(result[8][0])
		

	xven[idx][0] = timefrom
	idx += 1


x = np.array(xven)
y = np.array(trainvec)
#print(x)
#print(y)
X = np.transpose(x)
print(len(X))
print(len(X[1]))
print(len(y))
print(len(y[1]))
np.set_printoptions(threshold=np.nan)

# for h in xven:
# 	for t in h:
# 		print(str(t), end=" ")

# 	print("")

# for h in x:
# 	print(h)

# for h in np.transpose(y):
# 	print(h)

from keras.models import Sequential
from keras.layers import Dense
from sklearn.preprocessing import normalize

# xnorm = normalize(x, axis=1, norm='l1')
# ynorm = normalize(np.transpose(y), axis=1, norm='l1')

xnorm = x
ynorm = np.transpose(y)

model = Sequential()
model.add(Dense(86, input_dim=43, activation='tanh'))
model.add(Dense(60, activation='tanh'))
model.add(Dense(9, activation='linear'))
model.compile(loss='mse', optimizer='adam', metrics=['accuracy'])

model.fit(xnorm, ynorm, epochs=650, batch_size=50)
model.save('modelpred.h5')
scores = model.evaluate(xnorm, ynorm)
print("\n%s: %.2f%%" % (model.metrics_names[1], scores[1]*100))
# row1 = sheet.row(3)
# cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
# print(cell_type_str)
# if cell_type_str == "":
# 	print("banan")4,0
# print(type(cell_type_str))


xtest = []
xtest.append(0.5)
if str(data[3][1]) != "":
	xtest.append(float(str(data[3][1]).replace(",", ".")))
else:
	xtest.append(0.0)
if str(data[4][1]) != "":
	xtest.append(float(str(data[4][1]).replace(",", ".")))
else:
	xtest.append(0.0)
if str(data[5][1]) != "":
	xtest.append(float(str(data[5][1]).replace(",", ".")))
else:
	xtest.append(0.0)
if str(data[6][1]) != "":
	xtest.append(float(str(data[6][1]).replace(",", ".")))
else:
	xtest.append(0.0)
if str(data[7][1]) != "":
	xtest.append(float(str(data[7][1]).replace(",", ".")))
else:
	xtest.append(0.0)
if str(data[8][1]) != "":
	xtest.append(float(str(data[8][1]).replace(",", ".")))
else:
	xtest.append(0.0)
if str(data[9][1]) != "":
	xtest.append(float(str(data[9][1]).replace(",", ".")))
else:
	xtest.append(0.0)	# xven[i].append(float(str(data[3][i]).replace(",", ".")))	# xven[i].append(float(str(data[4][i]).replace(",", ".")))	# xven[i].append(float(str(data[5][i]).replace(",", ".")))	# xven[i].append(float(str(data[6][i]).replace(",", ".")))	# xven[i].append(float(str(data[7][i]).replace(",", ".")))	# xven[i].append(float(str(data[8][i]).replace(",", ".")))	# xven[i].append(float(str(data[9][i]).replace(",", ".")))if str(data[10][1]) != "":	xtest.append(float(str(data[10][1]).replace(",", ".")))else:	xtest.append(0.0)if str(data[11][1]) != "":	xtest.append(float(str(data[11][1]).replace(",", ".")))else:	xtest.append(0.0)if (data[15][1] != ""):	xtest.append(float(data[15][1]))else:	xtest.append(0.0)index = len(xtest)for j in range(0, 32):	#print(data[2][i] + " : " + str(companies.get(data[2][i])))		if j == companies[data[2][1]]:				xtest.append(1)		else:			xtest.append(0)



