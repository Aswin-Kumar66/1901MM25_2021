import re
import os
import shutil

def regex_renamer():

	# Taking input from the user

	print("1. Breaking Bad")
	print("2. Game of Thrones")
	print("3. Lucifer")

	webseries_num = int(input("Enter the number of the web series that you wish to rename. 1/2/3: "))
	season_padding = int(input("Enter the Season Number Padding: "))
	episode_padding = int(input("Enter the Episode Number Padding: "))
	if webseries_num == 1 :
         shutil.copytree(src ='./wrong_srt/Breaking Bad'  , dst= os.getcwd() + '/corrected_srt/Breaking Bad')
	elif webseries_num == 2 :
		 shutil.copytree(src ='./wrong_srt/Game of Thrones'  , dst= os.getcwd() + '/corrected_srt/Game of Thrones')
	elif webseries_num == 3 :
		 shutil.copytree(src ='./wrong_srt/Lucifer'  , dst= os.getcwd() + '/corrected_srt/Lucifer')

		 
   
	
	if ( webseries_num == 1) :
		os.chdir('./corrected_srt/Breaking Bad')		
	if webseries_num == 2 :
		os.chdir('./corrected_srt/Game of Thrones')
	if webseries_num == 3 :
		os.chdir('./corrected_srt/Lucifer')
		
		
	GT = os.listdir(os.getcwd())
	for name in GT :
		words = name.split('.')
		words = words[0] +'.' + words[-1]
		pattern = re.compile(r'\s'+'\d+' + '\w' + '\.')
		words = pattern.sub('.',words)
		os.rename(name , words)
		
			
		
	
	B = os.listdir(os.getcwd())
	for n in B :
			pattern1 = re.compile(r'.' +'\d+' + '[a-zA-Z]' + '\d+')
			pattern2 = re.compile(r'\d+')
			found = re.findall(pattern1 , n)
			found1 = re.findall(pattern2 , found[0])
			if found[0][0] == ' ' :
			 os.rename(n ,pattern1.sub(' Season ' + found1[0].zfill(season_padding) + ' Episode '  + found1[1].zfill(episode_padding)   , n ))
			else :
			 os.rename(n ,pattern1.sub('Season ' + found1[0].zfill(season_padding) + ' Episode '  + found1[1].zfill(episode_padding)   , n ))

	
	
	

regex_renamer()