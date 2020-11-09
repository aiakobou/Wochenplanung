#!/usr/bin/env python

__author__ = 'Andreas Iakobou'
__version__ = '2.0.0'
__email__ = 'andreas.iakobou@partner.bmw.de'

import os
try:
	os.chdir(os.path.dirname(__file__))
except:
	pass		
from tkinter import*
from tkinter import ttk
import tkinter as tk
from tkinter import messagebox as box
from .files import wochenplanung as wopla
from .files.calweek import calweek
from .files.copyxlsx import copyx
from .files.planning import planning
from .files.read_data import read_eBas
from .files.wochenplanung import wochenplanung
from .files.writexlsx import prepare_excel, write_data
import datetime


class MainGUI(Frame):
	
	def __init__(self, parent):		
		self.akfAnzahl = 0
		self.akfEingeplant = 0
		self.akfAussortiert = 0
		
		Frame.__init__(self, parent, background="light grey")
		self.parent = parent
		self.parent.title("AKF-Wochenplanung")
		self.parent.iconbitmap(r'imgs\icon.ico')
		self.style = ttk.Style()
		self.style.theme_use("default")
		self.centreWindow()
		self.pack(fill=BOTH, expand=1)

		menubar = Menu(self.parent)
		self.parent.config(menu=menubar)
		fileMenu = Menu(menubar)
		fileMenu.add_command(label="Info", command=self.info)
		fileMenu.add_command(label="Hilfe", command=self.help)
		fileMenu.add_command(label="Exit", command=self.quit)
		menubar.add_cascade(label="Menü", menu=fileMenu)
				
		PrüffeldLabel = Label(self, text="Prüffeld")
		PrüffeldLabel.grid(row=0, column=0, sticky=W+E)
		JahrLabel = Label(self, text="Jahr")
		JahrLabel.grid(row=1, column=0, sticky=W+E)
		KWLabel = Label(self, text="KW")
		KWLabel.grid(row=2, column=0, sticky=W+E)
		DateiLabel = Label(self, text="Datei auswählen")
		DateiLabel.grid(row=3, column=0, pady=10, sticky=W+E+N)
		StatusText = Label(self, text = 'Status: ', bg='light grey')
		StatusText.grid(row = 4, column = 2, padx=5, pady=5, ipady=2, sticky=W)	
		
		self.PrüffeldVar = StringVar()
		self.PrüffeldCombo = ttk.Combobox(self, textvariable=self.PrüffeldVar)
		self.PrüffeldCombo['values'] = ('EETZ', 'UPZ')
		self.PrüffeldCombo.current(1)
		self.PrüffeldCombo.grid(row=0, column=1, padx=5, pady=5, ipady=2, sticky=W)
		
		self.JahrVar = IntVar()
		self.JahrCombo = ttk.Combobox(self, textvariable=self.JahrVar)
		self.JahrCombo['values'] = (2020, 2021, 2022, 2023, 2024)
		self.JahrCombo.grid(row=1, column=1, padx=5, pady=5, ipady=2, sticky=W)
		self.JahrCombo.current(0)
		
		kws = []
		for i in range(1, 53):
			kws.append(i)
		self.KWVar = IntVar()
		self.KWCombo = ttk.Combobox(self, textvariable=self.KWVar)
		self.KWCombo['values'] = (kws)
		self.KWCombo.grid(row=2, column=1, padx=5, pady=5, ipady=2, sticky=W)
		self.KWCombo.current(calweek())
		
		filelist = self.getFileName()
		self.DateiVar = StringVar()
		self.DateiCombo = ttk.Combobox(self, textvariable=self.DateiVar)
		self.DateiCombo['values'] = (filelist)
		self.DateiCombo.grid(row=3, column=1, padx=5, pady=5, ipady=2, sticky=W)
		
		self.StatusText = Label(self, text = '   bereit   ', bg='green')
		self.StatusText.grid(row = 4, column = 3, padx=5, pady=5, ipady=2, sticky=W)	
				
		okBtn = Button(self, text="Wochenplanung erstellen", width=10, command=lambda: self.onConfirm())
		okBtn.grid(row=4, column=0, columnspan =2, padx=5, pady=3, sticky=W+E)
		closeBtn = Button(self, text="Close", width=10, command=self.onExit)
		closeBtn.grid(row=4, column=4, padx=5, pady=3, sticky=W+E)
		

	def onConfirm(self):		
		
		try:
			self.ResultText.grid_forget()
			self.ResultText.update()
		except:
			pass		
		self.StatusText.grid_forget()
		self.update_status(1)
		self.StatusText.update()
		
		try:
			if self.DateiVar.get() == '':
				box.showinfo("Fehler", "Keine Datei ausgewählt")
				self.StatusText.grid_forget()
				self.update_status(2)				
				self.StatusText.update()
			else:
				wopla.wochenplanung(self.PrüffeldVar, self.JahrVar, self.KWVar, self.DateiVar)
				self.StatusText.grid_forget()
				self.update_status(2)
				self.StatusText.update()
				self.ResultText = Label(self, text = 'Wochenplanung erstellt. Bitte Ergebnis überprüfen!', bg = 'light grey')
				self.ResultText.grid(row = 5, column = 0, columnspan =3, padx=5, pady=5, ipady=2, sticky=W)
		except Exception as e:
			self.StatusText.grid_forget()
			self.update_status(0)
			self.StatusText.update()
			raise(e)
			
					
	def update_status(self,v):		
		if v ==1:
			StatusText = Label(self, text = 'in Arbeit ', bg='yellow')
			StatusText.grid(row = 4, column = 3, padx=5, pady=5, ipady=2, sticky=W)	
		elif v ==0:
			StatusText = Label(self, text = '   Fehler   ', bg='red')
			StatusText.grid(row = 4, column = 3, padx=5, pady=5, ipady=2, sticky=W)	
		else:
			StatusText = Label(self, text = '   bereit   ', bg='green')
			StatusText.grid(row = 4, column = 3, padx=5, pady=5, ipady=2, sticky=W)	
			
			
	def info(self):
		box.showinfo("Information", "AKF-Wochenplanung für EA-8 AKFs \nVersion 2.0.0\nCAD Technik Kleinkoenen GmbH\nAndreas Iakobou - 28.05.2020 - \nandreas.iakobou@partner.bmw.de")
		
	
	def help(self):
		box.showinfo("Hilfe", "Die eBAS-Export-Datei muss sich im selben Ordner wie das Programm befinden.\nWenn ein Fehler auftritt, das Programm beenden und neu starten. Falls dies nicht hilft, bitte Info an Autor")
		
		
	def onExit(self):
		self.quit()
		

	def centreWindow(self):	
		w = 450
		h = 250
		sw = self.parent.winfo_screenwidth()
		sh = self.parent.winfo_screenheight()
		x = (sw - w)/2
		y = (sh - h)/2
		self.parent.geometry('%dx%d+%d+%d' % (w, h, x, y))
		
		
	def getFileName(self):
		filelist = list()
		for filename in os.listdir('..'):
			if filename.endswith(".xlsx") or filename.endswith(".xls"):
				filelist.append(filename)
			else:
				continue
		return filelist

	
def main():

	root = Tk()
	root.resizable(width=FALSE, height=FALSE)
	app = MainGUI(root)
	root.mainloop()
	

if __name__ == '__main__':
    main()

