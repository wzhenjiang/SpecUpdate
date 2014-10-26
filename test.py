#!/usr/bin/python
# -*- coding: utf-8 -*-

# ######################################################################
# This file is to help structuring a good program
# Created in 2013
# ######################################################################

from win32com.client import constants,Dispatch
import os,sys,datetime,time,shutil


# Below is the area for class definition

# ZCY

			
# Below is for helper functions

def main():

	# Here, your unit test code or main program
	
	# you can change below paras to define another unit test case
	
	last_spec = 'd:/test/2011_old.xls'
	new_spec = 'd:/test/2011_new.xls'
	
	
	xlsApp = Dispatch("Excel.Application")
	
	last_book = xlsApp.Workbooks.open(last_spec)
	print last_spec, 'loaded'
	new_book = xlsApp.Workbooks.open(new_spec)
	print new_spec, 'loaded'
	
	last_sheets = last_book.Sheets
	new_sheets = new_book.Sheets

	COL = 6
	
	for sheet in new_sheets:
		print sheet.Name
		if sheet.Name =='ZCY':
			sht = last_book.Worksheets(sheet.Name)
			for row_index in range(1,200):
				if (not sheet.Cells(row_index,COL).Value) and sht.Cells(row_index,COL).Value:
					sheet.Cells(row_index,COL).Value = sht.Cells(row_index,COL).Value
					print 'add value:', sht.Cells(row_index,COL).Value
		
	new_book.Save()
	new_book.Close()
	print new_spec, 'closed'
	last_book.Close(SaveChanges=0)
	print last_spec, 'closed'
	
	del xlsApp

if __name__=='__main__':
	main()