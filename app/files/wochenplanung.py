__author__ = 'Andreas Iakobou'
__version__ = '2.0.0'

import os
from datetime import date
from datetime import timedelta
import shutil
import openpyxl
from .calweek import calweek
from .planning import planning
from .read_data import read_eBas
from .copyxlsx import copyx
from .writexlsx import prepare_excel



def wochenplanung(PrüffeldVar = '', yearVar = 0, KWVar=0, fileVar=''):
	
	#In Version 2.0 for GUI, the steps for creating file and KW are no longer required but are kept if Wochenplanung is run as __main__ 
	if yearVar == 0:
		P = input('Prüffeld: ')
		if P in ['e', 'E', 'eetz', 'EETZ']:
			Prüffeld = 'EETZ'
		else:
			Prüffeld = 'UPZ'
		
		while True:
			KWstr = input('Bitte KW eingeben: ')
			yearstr = input('Jahr eingeben: ')
			
			try:
				assert len(yearstr) == 4
				KW = int(KWstr)
				year = int(yearstr)
				break
			except AssertionError as e:
				e.args += ('Ungültige Jahreszahl', yearstr)
				raise
			except:
				print('Falsches Format bei KW oder Jahr')
		
		while True:
			filename = input('Dateiname eingeben: ')
			file = r'..\{}'.format(filename)
			if os.path.isfile(file):
				break
			else:
				print('Datei nicht gefunden. Erneut eingeben.\n')
	else:
		Prüffeld = PrüffeldVar.get()
		year = yearVar.get()
		KW = KWVar.get()
		file = fileVar
	
	#Copy Excel Template and create new File
	fname = copyx(Prüffeld,KW)
	
	# Get Monday for current CW
	monday = date.fromisocalendar(year, KW, 1)

	# Open workbook
	workbook = openpyxl.load_workbook(fname)
	sheet = workbook.active
	
	# Update Excel with current date
	prepare_excel(monday, workbook, sheet, Prüffeld, KW, year)

	# Read the AKF data from eBas Export
	akfs = read_eBas(file, Prüffeld)

	#Call planning function
	planning(workbook, sheet, monday, akfs, Prüffeld)

	workbook.save(fname)
	workbook.close()
	shutil.move(fname, "..\{}".format(fname))

	
if __name__ == '__main__':
    wochenplanung()


