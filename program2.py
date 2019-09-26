import sys
from docx import Document # may need to install this package! 
from io import StringIO

input_file = sys.argv[1]

with open(input_file, 'rb') as f:
	doc = Document(input_file)
	fullText = []
	for para in doc.paragraphs:
		fullText.append(para.text)
	print('\n'.join(fullText))
	#source_stream = StringIO(f.read())