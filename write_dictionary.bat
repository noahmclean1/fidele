@ECHO OFF
ECHO Cloning GitHub repository ...
git clone https://github.com/noahmclean1/fidele.git
ECHO Entering GitHub repository ...
cd fidele\
ECHO Installing packages ...
python -m pip install python-docx --user
python -m pip install xlsxwriter
python -m pip install xlrd
python -m pip install xlwt
ECHO Running program 1 ...
python3 program1.py English_entries_with_DUPLICATES_WORD_LIST.docx redundantLines.docx
ECHO Opening Word file with redundant lines underlined ...
redundantLines.docx
ECHO Opening Excel file with frequency table ...
frequencyTable.xlsx
ECHO Running program 2 ...
python3 program2.py English_SERIES_alpha_TEXT.docx frequencyTable.xlsx
ECHO Opening Excel file with reference table ...
referenceTable.xlsx