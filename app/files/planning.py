#!/usr/bin/env python

'''Funktion zur Zuweisung der AKFs zu den Excel Zellen'''

__author__ = 'Andreas Iakobou'
__version__ = '1.0'
__email__ = 'andreas.iakobou@partner.bmw.de'

import datetime
from datetime import timedelta
from .writexlsx import write_data


def planning(workbook, sheet, monday, akfs, Prüffeld):

    # Create list with weekdates(Sun-Sat)
    days = list()
    days.append(monday.strftime("%d.%m.%Y"))
    for i in range(1,6):
        days.append((monday + timedelta(days=i)).strftime("%d.%m.%Y")) #Mon-Sat

    days.insert(0, (monday - timedelta(days=1)).strftime("%d.%m.%Y")) #Add Sunday


    #Assign AKFs on Sunday (not part of main loop)
    for akf in akfs:
        if akf[1].split()[0] == days[0]:
            startH = int(akf[1].split()[1].split(':')[0])
            endH = int(akf[2].split()[1].split(':')[0])
            roww = 6
            columnn = 4
            colorcode = 'b'

            if akf[2].split()[0] == days[1] and endH >= 11:
                roww = roww  + 2
                colorcode = 'ctk'
            elif akf[2].split()[0] == days[0]:
                roww = 3

            write_data(workbook, sheet, roww, columnn, akf, colorcode, Prüffeld)

    # Loop through all days (mon-Sat)
    rowoffset = 0
    for day in days[1:]:

        # Loop through all AKFs and assign to day
        for akf in akfs:

            startH = int(akf[1].split()[1].split(':')[0])
            endH = int(akf[2].split()[1].split(':')[0])
            roww = 6
            columnn = 4
            colorcode = 'ctk'


            if day !=days[0] and akf[1].split()[0] == day:
                if 0 <= startH <= 5:
                    if endH >= 11:
                        roww = roww + rowoffset + 2
                    else:
                        roww = roww + rowoffset
                        colorcode = 'b'

                elif 5 < startH <= 9:
                    roww = roww + rowoffset + 4
                elif 9 < startH <= 11:
                    roww = roww + rowoffset + 5
                elif 11 < startH <= 13:
                    roww = roww + rowoffset + 6
                elif 13 < startH <= 15:
                    roww = roww + rowoffset + 7
                elif 15 < startH <= 17:
                    if akf[2] != day and endH >= 11 and akf[1].split()[0] != days[5]:
                        roww = roww + rowoffset + 15
                    else:
                        roww = roww + rowoffset + 8
                        colorcode = 'b'
                elif 17 < startH <= 21:
                    if akf[2] != day and endH >= 11 and akf[1].split()[0] != days[5]:
                        roww = roww + rowoffset + 15
                    else:
                        roww = roww + rowoffset + 10
                        colorcode = 'b'
                else:
                    if akf[2] != day and endH >= 11 and akf[1].split()[0] != days[5]:
                        roww = roww + rowoffset + 15
                    else:
                        roww = roww + rowoffset + 13
                        colorcode = 'b'

                if roww > 69:
                    roww = 69
                    colorcode = 'b'

                write_data(workbook, sheet, roww, columnn, akf, colorcode, Prüffeld)
                #End if statement
            #End for akf in akfs loop

        rowoffset+=13
        #End for day in days loop

    #Handle unplanned akfs
    akfplanned = 0
    sorted_out = 0
    for akf in akfs:
        if len(akf) == 5:
            sorted_out+=1
            roww = 78
            columnn = 4
            colorcode = 'grey'
            write_data(workbook, sheet, roww, columnn, akf, colorcode, Prüffeld)
        else:
            akfplanned+=1

	
    #return(len(akfs), akfplanned, sorted_out)
