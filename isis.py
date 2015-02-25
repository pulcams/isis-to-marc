#!/usr/bin/env python
#-*- coding: utf-8 -*-

"""
Crude little script to convert spreadsheets from Turkish vendor Isis (http://www.theisispress.org/) into MaRC for loading into Voyager. 
It converts the xlsx to csv, then parses the csv. Requires MarcEdit.
Note: No need to notify KE when resultant files are put into load folder. 
Run like this: `python isis.py -f 2014-9_Prin_inv_no_221.xlsx -s 96,189 -i 221`
from 20141106
pmg
"""

import argparse
import glob
import os
import re
import subprocess
import shutil
import sys
import time
import unicodecsv
import xlrd

# argparse
parser = argparse.ArgumentParser(description='Process Isis spreadsheets.')
parser.add_argument('-f','--filename',type=str,dest="workbook",help="The full name of spreadsheet, e.g. '2014-7_inv_no_210_Prin.xlsx'",required=True)
parser.add_argument('-s','--split',type=str,dest="split",help="The line numbers *after which* to split the records. E.g. 96, 189",required=False)
parser.add_argument('-i','--invoice',type=str,dest="invoice",help="The invoice no. (not in csv).",required=False)
args = vars(parser.parse_args())

INDIR = './in/'
TEMPDIR = './temp/'
ARCHIVE = './archive/'
cmarcedit = '/opt/local/marcedit/cmarcedit.exe'
load = "/mnt/lib-tsserver/vendor_records_IN/Isis_input/"
today = time.strftime('%Y%m%d')
workbook = args['workbook']
invoiceno = args['invoice']
split = [1] # record numbers after which to split mrk file, provided via -s flag (start with row 1)
if args['split']:
	splits = args['split']
	for s in splits.split(','):
		split.append(s)
		
def setup():
	'''
	Just make sure there are in/ and out/ dirs
	'''
	if not os.path.isdir(INDIR):
		os.mkdir(INDIR,0775)
	if not os.path.isdir(TEMPDIR):
		os.mkdir(TEMPDIR,0775)
	if not os.path.isdir(ARCHIVE):
		os.mkdir(ARCHIVE,0775)
		
def csv_from_excel():
	'''
	Open workbook and write to csv file. 
	'''
	wb = xlrd.open_workbook(INDIR + workbook, encoding_override="cp1252")
	sh = wb.sheet_by_name('Sayfa1') # <= worksheet
	csv_file = open(INDIR + 'data.csv', 'w+b')
	wr = unicodecsv.writer(csv_file, quoting=unicodecsv.QUOTE_ALL, encoding='utf-8')
	
	for rownum in xrange(sh.nrows):
		a = list(x.encode('utf-8') if type(x) == type(u'') else x
				for x in sh.row_values(rownum))
		if (a[0] != '' and a[1] != ''): # ignore all the invisible / faulty rows
			wr.writerow(a)
	
	csv_file.close()

def data_from_csv():
	'''
	Grab data from csv file and write to mnemonic MaRC file
	'''
	mrko = re.sub('\.xlsx',"",workbook) # for naming out files
	mrk = mrko
	with open("./in/data.csv","rb") as csvfile:
		reader = unicodecsv.reader(csvfile,delimiter=',', quotechar='"', encoding='utf-8')
		next(reader, None)  # skip the headers
		rowcount = 1
		for row in reader:
			# Split invoice, if necessary. See the -s flag. If no -s is given (record numbers after which to split file), a single file is produced.
			breakpt = str(rowcount)
			if breakpt in split:
				mrk = mrko + '_' + breakpt
				
			with open(TEMPDIR+mrk + ".mrk","ab") as outfile:
				yr = re.sub('\.0',"",row[5]).encode('utf-8')
				if yr is None:
					yr = '\\\\'
				isbn = re.sub('\.0',"",row[7]).encode('utf-8')
				lineno = row[0]
				ti = row[1].encode('utf-8')
				au = row[2].encode('utf-8')
				place = row[3].encode('utf-8')
				pub = row[4].encode('utf-8')
				pp = row[6].encode('utf-8')
				price = row[8].encode('utf-8')
				if price != '':
					price = '{0:.2f}'.format(float(price))
				else: 
					price = '0.00'
				#print(price)
				# LDR
				outfile.write("=LDR  00000nam a2200000ia 4500"+"\r\n")
				# 008
				outfile.write("=008  140221s"+yr+"\\\\\\\\tu\\\\\\\\\\\\\\\\\\\\\\\\000\\0\\tur\\d"+"\r\n")
				# 020
				if isbn:
					outfile.write("=020  \\\\$a"+isbn+"\r\n")
				# 100
				if au:
					outfile.write("=100  1\$a"+au+"\r\n")
				# 245
				# TODO: non-filing characters
				ind2 = '0'
				if re.match('The', ti):
					ind2 = '4'
				if re.match('An', ti):
					ind2 = '3'
				outfile.write("=245  1"+ind2+"$a"+ti+"\r\n")
				# 260
				f260 = "=260  \\\\$a"+place
				if pub:
					f260 += " :$b"+pub
				if yr: 
					f260 += ",$c"+yr+"."
					#pass
				outfile.write(f260+"\r\n")
				# 300
				outfile.write("=300  \\\\$a"+pp+"\r\n")
				# 945
				if lineno:
					outfile.write("=945  \\\\$a"+lineno+"\r\n")
				# 980 ($e is the only important one)
				if invoiceno:
					outfile.write('=980  \\\\$e'+price+'$f'+invoiceno+'\r\n')
				else:
					outfile.write('=980  \\\\$e'+price+'\r\n')
				# the terminating carriage return
				outfile.write('\r\n')
			rowcount += 1
	print("done.")
	
def make_mrc():
	'''
	Use MarcEdit cli to output mrc records.
	'''
	for mrkfile in glob.glob(r'./temp/*.mrk'):
		mrcfile = str(os.path.splitext(mrkfile)[0])+'.mrc'
		try:
			breakc = subprocess.Popen(['mono',cmarcedit,'-s',mrkfile,'-d',mrcfile,'-make'])
			breakc.communicate()
			print('made mrc')
		except:
			etype,evalue,etraceback = sys.exc_info()
			print("problem making mrc. %s" % evalue)
			
def mv_marc():
	'''
	Archive processed files.
	'''
	dest = ARCHIVE+today
	if not os.path.isdir(dest):
		try:
			os.mkdir(dest,0775)
		except:
			etype,evalue,etraceback = sys.exc_info()
			print("problem creating dir. %s" % evalue)

	if not glob.glob(r'./temp/*.mrc'):
		print("no mrc??")
		exit

	for mrk in glob.glob(r'./temp/*.mrk'):
		try:
			shutil.move(mrk,dest)
			print(mrk + " archived")
		except: 
			etype,evalue,etraceback = sys.exc_info()
			print("error moving mrk files.  %s" % evalue)
			pass
			
	for mrc in glob.glob(r'./temp/*.mrc'):
		try:
			shutil.copyfile(mrc,load+str(os.path.basename(mrc))) # move to isis_input
			print(mrc + " moved to load folder")
		except:
			etype,evalue,etraceback = sys.exc_info()
			print("error copying mrc files to load folder. %s" % evalue)
		
		try:
			shutil.move(mrc,dest) # archive
			print(mrc + " archived")
		except: 
			etype,evalue,etraceback = sys.exc_info()
			print("error moving mrc files. %s" % evalue)
			pass
	
	shutil.rmtree(TEMPDIR)

if __name__ == "__main__":
	setup()
	csv_from_excel()
	data_from_csv()
	make_mrc()
	mv_marc()
	print('='*18)
	print('that\'s all folks! might want to check the mrc files in isis_input on lib-tsserver.')
	print('='*18)
