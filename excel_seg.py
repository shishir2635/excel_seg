import xlrd,xlwt

print('Excel Segregator v1.0 by Ted_Mosby')

branch_list = ['CSE','IT','EEE','ME','CE','MCA','IC','EE','EC']

#taking inputs from the user

for branch in branch_list:
	workbook = xlrd.open_workbook('Book1.xlsx')
	wb = xlwt.Workbook() 

	worksheet = workbook.sheet_by_index(0)
	ws = wb.add_sheet('{}'.format(branch))

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
		if cur_rec[1] == branch:  # change the value to any parameter column no. you want to match
			for y in range(len(cur_rec)):
				arow.append(cur_rec[y])

			filter_list.append(arow)
			arow = []

	if len(filter_list) == 0:
		continue

	for x in range(len(filter_list)):
		cur_stu = filter_list[x]
		for y in range(len(cur_stu)):
			ws.write(x+2,y,cur_stu[y])

	for z in range(len(filter_list[0])):  # write title of the fields
		ws.write(0,z, all_students[0][z])



	wb.save('{} - {} records.xls'.format(branch,len(filter_list)))
