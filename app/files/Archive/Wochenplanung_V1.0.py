import datetime
from datetime import date
from datetime import timedelta
import shutil
import openpyxl
from files.calweek import calweek
from files.planning import planning
import os
from files.read_data import read_eBas

file = input('Dateiname eingeben: ')
input('\nDie Datei eBas_Export: '+file+' in den Programmordner schieben.\n\nMit Enter bestätigen')


# Get current Calendar Week or desirede CW and monday of that week
year = input('\nJahr eingeben oder Enter für 2020: ')
if len(year)<1:
    year = '2020'
KW = calweek()

print('\n\nArbeite...')

# Copy Excel Template to create new Workbook, if file exists create Version 2, 3, ...
fname = "AKF-Wochenplanung_UPZ_KW_{strKW}.xlsx".format(str(KW))
#fname = 'AKF-Wochenplanung_UPZ_KW_'+ str(KW) + '.xlsx'
i = 1
while True:
    if os.path.isfile(fname) == False:
        shutil.copyfile('files/Template.xlsx',fname)
        break
    else:
        i+=1
        fname = "AKF-Wochenplanung_UPZ_KW_{strKW}_V{strVersion}.xlsx".format(str(KW),str(i))
        #fname = 'AKF-Wochenplanung_UPZ_KW_'+ str(KW) + '_V' + str(i) + '.xlsx'


# Get Monday for current CW
monday = date.fromisocalendar(int(year), KW, 1)

# Open workbook
workbook = openpyxl.load_workbook(fname)
sheet = workbook.active

# Write CW and monday date to sheet
sheet['D1'] = 'KW ' + str(KW)
sheet['A6'] = monday.strftime("%d.%m.%Y")

# Read the AKF data from eBas Export
akfs = read_eBas(file)

#Call planning function
planning(workbook, sheet, monday, akfs)


workbook.save(fname)
workbook.close()

input('\nExcel-Datei mit dem Namen ' + fname + ' erstellt. Enter zum Beenden drücken')
