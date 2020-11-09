#!/usr/bin/env python

'''Einlesen der eBAS Datei'''

__author__ = 'Andreas Iakobou'
__version__ = '1.0'
__email__ = 'andreas.iakobou@partner.bmw.de'

import os

def read_eBas(fileVar, Prüffeld):
	import xlrd
	from operator import itemgetter

	# Handle file formats. Could be StringVar or str

	try:
		file = fileVar.get()
		path = r"..\{}".format(file)
		
	except:
		file = fileVar
		path = file
	# Open the file and select sheet

	book = xlrd.open_workbook(path)
	sheet1 = book.sheet_by_index(0)

	# Create the nested list of AKFs
	akf_list_upz = list()
	akf_list_eetz = list()
	for i in range(len(sheet1.col_values(0))):   # Go through all rows except row 0
		akf = list()
		if i == 0:
			continue
		else:
			row = sheet1.row_values(i)
			akf.extend(row[j] for j in [3,5,6,7,16])
			if akf[0] == 'UPZ - E':
				akf_list_upz.append(akf)
			if akf[0] == 'EETZ - E':
				akf_list_eetz.append(akf)

	#Sort by date
	if Prüffeld == 'UPZ':
		akf_list_sorted = sorted(akf_list_upz, key=itemgetter(2))
	if Prüffeld == 'EETZ':
		akf_list_sorted = sorted(akf_list_eetz, key=itemgetter(2))

	return(akf_list_sorted)
