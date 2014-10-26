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

def print_help():
	print 'Purpose: udpate spec with value from last version'
	print 'Command line: ./SpecUpdate.py last_spec_path new_spec_path'


#Below is main function
def main():

	# Here, your unit test code or main program
	
	# you can change below paras to define another unit test case

	print 'SpecUpdate 1.0'
	print '-----------------------------------------------------------'
	
	argc = len(sys.argv)
	if argc <= 2:
		print_help()
		exit()

	last_spec_name = sys.argv[1]
	new_spec_name = sys.argv[2]
	
	# below two lines are used for test only
	last_spec_name = 'd:/LocalDev/Project/specupdate/MGCF-PM(V2.1.5)-Company-Version00.xlsx'
	new_spec_name = 'd:/LocalDev/Project/specupdate/MGCF-PM(V2.1.6)-Company-Version00.xlsx'
	
	# Generate xlsApp to invoke win32com
	xlsApp = Dispatch("Excel.Application")
	
	#load workbooks
	print 'Load excel files ...'
	last_book = xlsApp.Workbooks.open(last_spec_name)
	print last_spec_name, 'loaded'
	new_book = xlsApp.Workbooks.open(new_spec_name)
	print new_spec_name, 'loaded'
	
	#Prepare sheets to compare with
	last_sheets = last_book.Sheets
	new_sheets = new_book.Sheets

	#select column to compare with
	key_col = 1		#A Column
	value_col = 13	#M Column
	max_rows = 5000
	threshold_none = 5
	
	#define start sheet
	sheet_start = 3 	#skip sheet 0 and sheet 1
	
	
	#build dict from old sheet
	print 'Build dictionary ...'
	last_sheet_dict = {}
	num_sheets = len(last_sheets)
	for sheet_index in range(sheet_start,num_sheets):
		sheet = last_sheets[sheet_index]
		sheet_name = sheet.Name
		print 'parsing sheet:', sheet_name
		value_dict = {}
		threshold_count = 0
		for row_index in range(2,max_rows):
			key = sheet.Cells(row_index,key_col).Value
			value = sheet.Cells(row_index,value_col).Value
			if key == None:
				threshold_count += 1
			if threshold_count > threshold_none:
				break
			if not (key == None) and (not value == None):
				value_dict[key] = value
				print key,':',value,'\t',
		if len(value_dict)> 0:
			print
		last_sheet_dict[sheet_name] = value_dict
	print 'Dictionary built.'
	
	#go through new sheet and fill in value
	print 'Filling new spec ...'
	num_sheets = len(new_sheets)
	for sheet_index in range(sheet_start,num_sheets):
		sheet = new_sheets[sheet_index]
		sheet_name = sheet.Name
		print 'Processing sheet:', sheet_name
		threshold_count = 0
		for row_index in range(2,max_rows):
			key = sheet.Cells(row_index,key_col).Value
			value = sheet.Cells(row_index,value_col).Value
			if key == None:
				threshold_count += 1
			if threshold_count > threshold_none:
				break
			if not (key == None) and (value == None):
				if sheet_name in last_sheet_dict.keys():
					if key in last_sheet_dict[sheet_name].keys():
						sheet.Cells(row_index,value_col).Value = last_sheet_dict[sheet_name][key]
						print sheet_name, key, 'set'
	print '\nNew spec is set.'
		
	new_book.Save()
	new_book.Close()
	print new_spec_name, 'closed'
	last_book.Close(SaveChanges=0)
	print last_spec_name, 'closed'
	
	del xlsApp
	
	print 'Everything cleared. Enjoy!'

if __name__=='__main__':
	main()