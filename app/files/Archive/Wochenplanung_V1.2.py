import datetime
from datetime import date
from datetime import timedelta
import shutil
import openpyxl
from files.calweek import calweek
from files.planning import planning
import os
from files.read_data import read_eBas
from files.copyx import copyx

Jahr = '2020'


def __init__(Prüffeld = 'UPZ'):

	'''Wochenplanung für EA-8 AKF Beladungen.'''
	
	P = input('Prüffeld auswählen E für EETZ, U für UPZ: ')
	if P in ['e', 'E', 'EETZ', 'eetz']:
		Prüffeld = 'EETZ'
	
	while True:
		filen = input('Dateiname eingeben: ')
		#if len(filen) < 1: filen = 'ebas_export.xlsx'
		file = 'AKF-Wochenplanung\{}'.format(filen)
		if os.path.isfile(file):
			break
		elif filen in ['x', 'X']:
			print('Programm beendet')	
			exit()
		else:
			print('Datei nicht gefunden. Erneut eingeben.\n Hinweis: Auf Endung ".xlsx" achten!\n "x" zum Beenden eingeben\n\n')

	# Get current Calendar Week or desirede CW
	year = input('\nJahr eingeben oder Enter für 2020: ')
	if len(year)<1:
		year = Jahr
	KW = calweek()

	print('\n\nArbeite...')

	#Copy Excel Template and create new File
	fname = copyx(Prüffeld,KW)
	
	# Get Monday for current CW
	monday = date.fromisocalendar(int(year), KW, 1)

	# Open workbook
	workbook = openpyxl.load_workbook(fname)
	sheet = workbook.active


	# NEW in V1.2: For Loop for creating list for weekdays and writing in Excel
	week = list()
	for i in range(-1, 6):
		week.append(monday+timedelta(days=i))

	if Prüffeld == 'EETZ':
		sheet['A1'] = 'KW{}'.format(str(KW))
		sheet['A2'] = '{}'.format(year)
	if Prüffeld == 'UPZ':
		sheet['D1'] = 'KW{}'.format(str(KW))
		sheet['A1'] = '{}'.format(year)
	for i, j in zip(['A4','A6','A19','A32','A45','A58','A71'], week):
		sheet[i] = j.strftime("%d.%m")

	# Read the AKF data from eBas Export
	akfs = read_eBas(file, Prüffeld)

	#Call planning function
	planning(workbook, sheet, monday, akfs, Prüffeld)


	workbook.save(fname)
	workbook.close()
	shutil.move(fname, "AKF-Wochenplanung\{}".format(fname))
	input('\nExcel-Datei mit dem Namen ' + fname + ' erstellt. Enter zum Beenden drücken')
	

if __name__ == '__main__':
    __init__()


