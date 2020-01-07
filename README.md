# fidele
Code for Swahili Dictionary Project

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
