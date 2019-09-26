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

# put lines of text into an array called lines 
num_lines = len(lines)
i = 0
while i != num_lines - 1:
    # number of occurrences of lines[i] = 1
    line = lines[i]
    j = i + 1
    while j != num_lines - 1:
        if lines[j] == line:
            # underline lines[j]
            # increment number of occurrences of that line
        j++
    i++
