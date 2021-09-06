

def output_individual_roll():
    
 import os
 
 if(not os.path.exists("output_individual_roll")):
    os.mkdir("output_individual_roll")
 

 with open('regtable_old.csv','r') as f:
     
     
     for line in f:
        words = line.split(',')
        if(words[0] == 'rollno') :
           head = line.split(',')
        
        if(words[0] != 'rollno') :
          rollno = '{}.csv'.format(words[0])
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
          

      
output_by_subject()
output_individual_roll()


