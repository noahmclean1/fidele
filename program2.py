import sys
from docx import Document # may need to install this package! 
import xlrd
import xlwt 
import xlsxwriter
#CREATE EXCEL FILE CALLED test.xlsx
workbook = xlsxwriter.Workbook('test.xlsx') 
worksheet = workbook.add_worksheet()
underline = workbook.add_format({'underline': True})

input_file = sys.argv[1]
freq_file = sys.argv[2]

'''
with open(input_file, 'rb') as f:
	doc = Document(input_file)
	fullText = []
	for para in doc.paragraphs:
		fullText.append(para.text)
	print('\n'.join(fullText))
'''
# Create Excel file to write into

# To read/write into excel files, look at the following snippet:

'''
# Reading an excel file using Python 
import xlrd 
  
# Give the location of the file 
loc = ("path of file") 
  
# To open Workbook 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
  
# For row 0 and column 0 
sheet.cell_value(0, 0).encode("ascii") 
'''

'''
Underlining parts of cells in Excel

import xlsxwriter
workbook = xlsxwriter.Workbook('test.xlsx')
worksheet = workbook.add_worksheet()
underline = workbook.add_format({'underline': True})
worksheet.write_rich_string(0, 0, 'under', underline, 'line') # 'line' is underlined, 'under' is not.
worksheet.write_rich_string('A2', underline, 'under', 'line') # 'under' is underlined, 'line' is not.
workbook.close()
'''


# Open the files
unitListFile = open(input_file, 'rb')
ulf = Document(unitListFile)
unitlist = []
for para in ulf.paragraphs:
	unitlist.append(para.text)

freqListFile = open(freq_file, 'rb')
flf = Document(freqListFile)
freqlist = []
for para in flf.paragraphs:
	freqlist.append(para.text)

# Create Excel file to write into

# For each line:
for units in unitlist:
	# Tokenize the units string
	toks = units.split(",")

	# Search for most frequent unit (down frequency list) -> dubbed B
	topword = None
	for freqword in freqlist:
		for token in toks:
			if freqword == token:
				topword = token
				break
		if topword != None:
			break

	# Check if none of the words were found, if so, pick the first (random)
	if topword == None:
		topword = toks[0]

	# Remove B from the read line, sort the rest
	rest = toks.remove(topword)
	rest = rest.sort()

	# Write B,[REST OF LINE] with B underlined into XLSX
	# Write B again, normally, into the right column

# def mainstream(word_list, freq_list):
    



workbook.save()
workbook.close()


