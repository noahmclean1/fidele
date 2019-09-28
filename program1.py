import sys
from docx import Document # may need to install this package! 
import operator

# Select file
input_file = sys.argv[1]


# Read in the full file contents into duplist
duplicateFile = open(input_file, 'rb')
ulf = Document(duplicateFile)
duplist = []
for para in ulf.paragraphs:
    duplist.append(para.text)

# Initialize a dictionary for frequency counts
freqCount = {}

# Find frequencies and underline redundancies
for i,phrase in enumerate(duplist):
    # First element
    if i == 0:
        freqCount[phrase] = 1
    else:
        if phrase in freqCount:
            freqCount[phrase] += 1
            # Is the line the same as the one above? (redundant)
            if phrase == duplist[i-1]:
                # UNDERLINE IN DOCX
                pass
        else:
            freqCount[phrase] = 1

sortFreq = sorted(freqCount.items(),key = operator.itemgetter(1), reverse=True)

for key,val in sortFreq:
    if val > 5:
        print("{} | {}".format(key,val))
    else:
        break