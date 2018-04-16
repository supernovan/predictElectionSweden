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



companies = {}
for x in data[2]:
	if x not in companies:
		companies[x] = len(companies)

xtest = []
xtest.append([])
xtest.append([])
for i in range(0, 2):
	xtest[i].append(0.5)
	if str(data[3][i]) != "":
		xtest[i].append(float(str(data[3][i]).replace(",", ".")))
	else:
		xtest[i].append(0.0)
	if str(data[4][i]) != "":
		xtest[i].append(float(str(data[4][i]).replace(",", ".")))
	else:
		xtest[i].append(0.0)
	if str(data[5][i]) != "":
		xtest[i].append(float(str(data[5][i]).replace(",", ".")))
	else:
		xtest[i].append(0.0)
	if str(data[6][i]) != "":
		xtest[i].append(float(str(data[6][i]).replace(",", ".")))
	else:
		xtest[i].append(0.0)
	if str(data[7][i]) != "":
		xtest[i].append(float(str(data[7][i]).replace(",", ".")))
	else:
		xtest[i].append(0.0)
	if str(data[8][i]) != "":
		xtest[i].append(float(str(data[8][i]).replace(",", ".")))
	else:
		xtest[i].append(0.0)
	if str(data[9][i]) != "":
		xtest[i].append(float(str(data[9][i]).replace(",", ".")))
	else:
		xtest[i].append(0.0)	# xven[i].append(float(str(data[3][ii).replace(",", ".")))	# xven[i].append(float(str(data[4][ii).replace(",", ".")))	# xven[i].append(float(str(data[5][ii).replace(",", ".")))	# xven[i].append(float(str(data[6][ii).replace(",", ".")))	# xven[i].append(float(str(data[7][ii).replace(",", ".")))	# xven[i].append(float(str(data[8][ii).replace(",", ".")))	# xven[i].append(float(str(data[9][ii).replace(",", ".")))if str(data[10][i]) != "":	xtest[i].append(float(str(data[10][i]).replace(",", ".")))else:	xtest[i].append(0.0)if str(data[11][i]) != "":	xtest[i].append(float(str(data[11][i]).replace(",", ".")))else:	xtest[i].append(0.0)if (data[15][i] != ""):	xtest[i].append(float(data[15][i]))else:	xtest[i].append(0.0)index = len(xtest[i])for j in range(0, 32):	#print(data[2][ii + " : " + str(companies.get(data[2][ii)))		if j == companies[data[2][1i]:				xtest[i].append(1)		else:			xtest[i].append(0)
	if str(data[10][i]) != "":
		xtest[i].append(float(str(data[10][i]).replace(",", ".")))
	else:
		xtest[i].append(0.0)
	if str(data[11][i]) != "":
		xtest[i].append(float(str(data[11][i]).replace(",", ".")))
	else:
		xtest[i].append(0.0)
		
	if data[15][i] != "":
		xtest[i].append(float(data[15][i]))
	else:
		xtest[i].append(0.0)
	for j in range(0, 32):
		#print(data[2][ii + " : " + str(companies.get(data[2][ii)))
		if j == companies[data[2][i]]:
			xtest[i].append(1)
		else:
			xtest[i].append(0)


xfinal = np.array(xtest)
print(xfinal)
print(len(xfinal))


from keras.models import load_model
model = load_model('modelpred.h5')

yhat = model.predict(xfinal, verbose=0)


print("M : " + str(yhat[0][0]))
print("L : " + str(yhat[0][1]))
print("C : " + str(yhat[0][2]))
print("KD : " +str( yhat[0][3]))
print("S : " + str(yhat[0][4]))
print("V : " + str(yhat[0][5]))
print("Mp : " +str( yhat[0][6]))
print("SD : " +str( yhat[0][7]))
print("Ö : " + str(yhat[0][8]))
