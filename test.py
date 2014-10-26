#!/usr/bin/python
# -*- coding: utf-8 -*-

# ######################################################################
# This file is to help structuring a good program
# Created in 2013
# ######################################################################

from win32com.client import constants,Dispatch
import os,sys,datetime,time,shutil


# Below is the area for class definition


			
# Below is for helper functions

def main():

	# Here, your unit test code or main program
	
	# you can change below paras to define another unit test case
	
	old_spec = 'd:/test/2011_old.xls'
	new_spec = 'd:/test/2011_new.xls'
	
	
	xlsApp = Dispatch("Excel.Application")
	
	old_book = xlsApp.Workbooks.open(old_spec)
	
	sheets = old_book.Sheets
	
	for sheet in sheets:
		print sheet.Name
		data = sheet.Range(sheet.Cells(1,1),sheet.Cells(500,50))
		for item in data:
			if item.Value:
				print item.Value,
		print ''

if __name__=='__main__':
	main()