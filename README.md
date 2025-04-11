**Auto-Translator**
This is a Python-based tool used for automatic translation and editing of Excel spreadsheet entries using a translator API. This program can iterate through and translate all the files in a given folder if they are formatted the same way and require the same parameters.

**Installation**
1. Clone the repository at https://github.com/Central-Asian-Archaeological-Landscapes/autotranslator.git and navigate to the project directory
2. Install dependencies via the requirements.txt file
3. Make sure the 'glossary.xlsx' file is formatted correctly and has all the entries that are required, ex: additional phrases that should be subbed in.

**Usage**
1. Data input: the first part of the code involves user input to set the parameters.
2. Enter the data path for the folder to be translated.
3. Enter input / output languages - current options are Russian and English, but the deeptranslate API supports numerous others - adding other languages will require some editing of code to add further language options
5. Enter the data path for the glossary file, as this may be in a different location - or if you are using a different glossary than the one provided.
6. The program will output what folder it is translating and all the files within that folder and matching subfolders
7. Remove any files that do not need to be translated or don't match the format by typing in the corresponding numbers, separated by comma - if all files are to be translated enter N
8. Enter which sheet needs to be translated within the spreadsheet (necessary for excel sheets with multiple tabs)
9. Enter the input and output columns. Multiple sets of columns can be translated for all rows, but they need to be separated by commas and be in the same order for both input/output. Data from two or more columns can be combined for input with a /. The same column can be edited with appended translated data if it is both input and output.
10. Enter the start row from which the program should operate, excluding column names
11. The program will output these parameters to confirm - once confirmed, the program will begin translation
12. Translation proceeds through several steps
    a. Entries from the relevant input columns are extracted by row, and stored as a set
    b. Each entry is compared with the glossary for matching phrases/words which are replaced using fnmatch
    c. The checked entry is fed through the deeptranslator API
    d. The translated entry is added into a new set
    e. The new set is input back into the Excel spreadsheet by row
    f. Where data needs to be combined from multiple columns, or input back into the same column, different sets are combined before inputting
    d. The program iterates through each Excel file to repeat the process

**Common Issues**
1. Make sure all files within the same folder/batch are formatted the same, and have the same columns to be translated/input. If not, the files in question can be removed during the load_data stage and translated separately.
2. Ensure all files that need to be translated as well as the glossary file are closed before use, as the program will encounter a permissions error if a file is open when it tries to access it.
3. Check that the correct sheet and columns have been selected
