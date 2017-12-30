import xlrd,xlwt

print('Excel Segregator v1.0 by Ted_Mosby')

workbook = xlrd.open_workbook('Book1.xlsx')
wb = xlwt.Workbook() 

worksheet = workbook.sheet_by_index(0)
ws = wb.add_sheet('EEE')

arow = list()
all_students = list()

# holding all the data in the record table
for x in range(worksheet.nrows):
	for y in range(worksheet.ncols):
		arow.append(worksheet.cell(x,y).value)
	all_students.append(arow)
	arow = []

filter_list = list()

for x in range(len(all_students)):
	cur_rec = all_students[x]
	if cur_rec[1] == 'EEE':
		for y in range(len(cur_rec)):
			arow.append(cur_rec[y])

		filter_list.append(arow)
		arow = []

for x in range(len(filter_list)):
	cur_stu = filter_list[x]
	for y in range(len(cur_stu)):
		ws.write(x,y,cur_stu[y])

wb.save('CSE.xls')
