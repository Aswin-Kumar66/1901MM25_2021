import os
from os import listdir
import os.path
from os.path import isfile, join
import csv
from openpyxl import Workbook

def output_individual_roll():
 from openpyxl import Workbook
 import os
 import csv
 
 if(not os.path.exists("output_individual_roll")):
    os.mkdir("output_individual_roll")
 

 with open('regtable_old.csv','r') as f:
     
     
     for line in f:
        words = line.split(',')
        if(words[0] == 'rollno') :
           head = line.split(',')
        
        if(words[0] != 'rollno') :
          rollno = '{}.xlsx'.format(words[0])
          path = "./output_individual_roll/"+ rollno
          if(os.path.isfile(path)):
             with open(path, 'a') as jfile:
             
             
              jfile.write(words[0])
              jfile.write(',')
              jfile.write(words[1])
              jfile.write(',')
              jfile.write(words[3])
              jfile.write(',')
              jfile.write(words[-1])
              jfile.close()
          else :
             with open(path, 'w') as jfile:
             
             
              jfile.write(head[0])
              jfile.write(',')
              jfile.write(head[1])
              jfile.write(',')
              jfile.write(head[3])
              jfile.write(',')
              jfile.write(head[-1])
              jfile.close()
          

          
     
        



def output_by_subject():
 import os
 
 if(not os.path.exists("output_by_subject")):
    os.mkdir("output_by_subject")
 

 with open('regtable_old.csv','r') as f:
     
     
     for line in f:
        words = line.split(',')
        
        if(words[3] == 'subno') :
           head = line.split(',')
        
        if(words[3] != 'subno') :
          subno = '{}.csv'.format(words[3])
          path = "./output_by_subject/"+ subno
          if(os.path.isfile(path)):
             with open(path, 'a') as jfile:
             
              
              jfile.write(words[0])
              jfile.write(',')
              jfile.write(words[1])
              jfile.write(',')
              jfile.write(words[3])
              jfile.write(',')
              jfile.write(words[-1])
              jfile.close()
          else :
             with open(path, 'w') as jfile:
             
              
              jfile.write(head[0])
              jfile.write(',')
              jfile.write(head[1])
              jfile.write(',')
              jfile.write(head[3])
              jfile.write(',')
              jfile.write(head[-1])
              jfile.close()
          

      


output_individual_roll()
output_by_subject()



def xlsx_output_by_rollno():
  files = [f for f in listdir('./output_individual_roll/') if isfile(join('./output_individual_roll/', f))]
  for file in files:
    if os.path.splitext(file)[1][1:] == 'csv':
      wb = Workbook()
      ws = wb.active

      with open('./output_individual_roll/' + file) as f:
          reader = csv.reader(f, delimiter=',')
          for row in reader:
              ws.append(row)
      
      file_name = os.path.splitext(file)[0]
      wb.save('./output_individual_roll/' + file_name + '.xlsx')
      os.remove('./output_individual_roll/' + file)


def xlsx_output_by_subject():
  files = [f for f in listdir('./output_by_subject/') if isfile(join('./output_by_subject/', f))]
  for file in files:
    if os.path.splitext(file)[1][1:] == 'csv':
      wb = Workbook()
      ws = wb.active

      with open('./output_by_subject/' + file) as f:
          reader = csv.reader(f, delimiter=',')
          for row in reader:
              ws.append(row)
      
      file_name = os.path.splitext(file)[0]
      wb.save('./output_by_subject/' + file_name + '.xlsx')
      os.remove('./output_by_subject/' + file)

xlsx_output_by_rollno()
xlsx_output_by_subject()


