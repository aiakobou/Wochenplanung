#!/usr/bin/env python

'''Berechnung der Kalenderwoche'''

__author__ = 'Andreas Iakobou'
__version__ = '1.0'
__email__ = 'andreas.iakobou@partner.bmw.de'

import datetime

def calweek():

    while True:
        KWalternativ = input('KW eingeben oder Enter für folgende KW: ')
        if len(KWalternativ) < 1:
            KW = datetime.date.today().isocalendar()[1] +1
            break
        else:
            try:
                KW = int(KWalternativ)
                break
            except:
                print('Keine gültige KW. Erneut eingeben')
    return(KW)
