# Copy Excel Template to create new Workbook, if file exists create Version 2, 3, ...

import os
import shutil

def copyx(Prüffeld,KW):


	fname = "{Prüf}_AKF-Wochenplanung_KW{strKW}.xlsx".format(Prüf=Prüffeld, strKW=KW)

	i = 1
	while True:
		if os.path.isfile(r'..\{}'.format(fname)) == False:
			shutil.copyfile(r'.\templates\Template_{}.xlsx'.format(Prüffeld), fname)
			break
		else:
			i+=1
			fname = "{Prüf}_AKF-Wochenplanung_KW{strKW}_V{strVersion}.xlsx".format(Prüf=Prüffeld,strKW=KW,strVersion=i)
				

			
	return fname


	
