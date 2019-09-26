import sys
from docx import Document # may need to install this package! 
import csv

input_file = sys.argv[1]

with open(input_file, 'rb') as f:
	doc = Document(input_file)
	fullText = []
	for para in doc.paragraphs:
		fullText.append(para.text)
	print('\n'.join(fullText))

# To read/write into CSV files, use csv.reader / csv.writer