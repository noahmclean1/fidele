import sys
from docx import Document # may need to install this package! 
import xlrd

input_file = sys.argv[1]

with open(input_file, 'rb') as f:
	doc = Document(input_file)
	fullText = []
	for para in doc.paragraphs:
		fullText.append(para.text)
	print('\n'.join(fullText))

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