#!/usr/bin/env python

'''Berechnung der Kalenderwoche'''

__author__ = 'Andreas Iakobou'
__version__ = '1.0'
__email__ = 'andreas.iakobou@partner.bmw.de'

import datetime

def calweek():
    
	KW = datetime.date.today().isocalendar()[1]

	   
	return(KW)
