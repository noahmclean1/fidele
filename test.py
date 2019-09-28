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

# UNIT TESTS
# Program 1

# Check that redundant lines are underlined
# Eyeball test: Choose a random page and ensure that 
#   - the words on the original doc and the generated doc match up
#   - redundant lines are underlined

# Check frequency numbers are correct
# 'abandon' (4), 'Bless you!' (2), 'break down' (3), 'breathe one’s last' (1), 'breath-taking' (1), 'coloring leaves or barks used to strengthen fishing lines' (1), 'East' (1), 'aire' (0)

# Check ordering of word list is correct
# 'start' (9), 'supply' (4), 'scrutinize' (2), 'superstitous person' (1) 

# Program 2
''' For each, we check lines with one unit and with multiple units
'''

# Check that each line has only one underlined unit

# Check that the underlined unit is the most frequent unit

# Check that the first unit is underlined

# Check that the rest of the units are in alphabetical order

# Check that the reference unit is repeated in column B
