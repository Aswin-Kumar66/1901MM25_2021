import csv
from openpyxl import Workbook
def feedback_not_submitted():

	data = ['roll' , 'registered_sem' ,'scheduled_sem ' , 'subno','Name', 'email','aemail' , 'contact']
	ltp_mapping_feedback_type = {1: 'lecture', 2: 'tutorial', 3:'practical'}
	output_file_name = "course_feedback_remaining.xlsx" 
	with open('course_registered_by_all_students.csv' , 'r') as f:
          course_registered = list(csv.reader(f))
	with open('course_feedback_submitted_by_students.csv' , 'r') as h :
          course_feedback  = list(csv.reader(h))
	with open('studentinfo.csv' , 'r') as i :
          studentinfo  = list(csv.reader(i))
	with open('course_master_dont_open_in_excel.csv' , 'r') as g :
          course_master  = list(csv.reader(g))
	wb = Workbook()
	sheet = wb.active
	sheet.append(data)

	for i in course_registered[1:] :
		append = []
		roll = i[0]
		subno = i[3]
		regsitersem = i[1]
		schedulesem = i[2]
		for course in course_master[1:] :
			count = 0
			if subno == course[0] :
				nonzero = 0
				ltp = course[2]
				splitltp = ltp.split('-')
				splitltp = [int(i) for i in splitltp]
				for i in splitltp :
					if i != 0 :
						nonzero = nonzero + 1
				for feedback in course_feedback :
                
			            if roll == feedback[3] :
						           
			    	         if subno == feedback[4] :
					             count = count + 1
				if count < nonzero :
					append.append(roll)
					append.append(regsitersem)
					append.append(schedulesem)
					append.append(subno)
					for name in studentinfo :
						if roll == name[1] :
						  append.append(name[0])
						  append.append(name[8])
						  append.append(name[9])
						  append.append(name[10])
					if len(append) < 8:
						for i in range(8-len(append)) :
							append.append('NA_IN_STUDENTINFO')
						
				
				if append != [] :
					sheet.append(append)
	wb.save(output_file_name)

			
		                        
	


 



feedback_not_submitted()
