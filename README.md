# fidele
Code for Swahili Dictionary Project

# HOW TO RUN **write_dictionary.bat**:
1. open **write_dictionary.bat**
2. copy all the text in the file
3. open **Notepad** on your computer. you can find it by typing "notepad" into the search bar at the bottom left and 
4. paste the text from **write_dictionary.bat** into Notepad
5. click **File** in the top left
6. click **Save As**
7. where it says **File name:** at the bottom, type "write_dictionary.bat"
8. click **Save** on the bottom left
9. go to where you saved the file and double-click it

# If you are working in Terminal:
Setup:
```
git clone https://github.com/noahmclean1/fidele.git
cd fidele\
pip3 install python-docx --user
pip3 install xlsxwriter --user
```

Program 1:
```
python3 program1.py English_entries_with_DUPLICATES_WORD_LIST.docx redundantLines.docx
open redundantLines.docx
open frequencyTable.docx
```

Program 2:
```
pip3 install xlrd --user
pip3 install xlwt --user
python3 program2.py English_SERIES_alpha_TEXT.docx frequencyTable.xlsx
open referenceTable.xlsx
```

To get permissions to execute shell file:
```
chmod u+x write_dictionary.sh
```
SOURCE: https://stackoverflow.com/questions/17015449/how-do-i-run-sh-or-bat-files-from-terminal
