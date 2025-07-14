'''
Created on 18 Dec 2021

@author: turch
'''


from translate_excel_to_csv import TranslateRun
import os
import pandas as pd
import traceback
import threading
import sys
from pathlib import Path


class LoadFolders: #creates a class to load folders within the given path for translation of all files within that folder
    archive = [] #creates list for archive spreadsheets
    monument = [] #and list for monument ones
    docs = []
    biglist = [archive, monument] #biglist for the two above lists so they can be cycled through
    l = None
    
    def sortfolder(self, data_path, filetype, folder):
        p = Path(data_path)
        if '.xl' in p.name:
            folder.append(str(p))
        else:
            for file_path in p.rglob('*'):
                if file_path.is_file():
                    folder.append(str(file_path))

    def openfolder(self, data_path, filetype): 
        '''
        This function is the first to be executed in the program as it locates the folders from the specified data path, accesses the files within them, and extracts the file paths to the specified lists (optional ultimately
        but useful for this group of spreadsheets). The file paths can then be used to access individual files for translation.
        '''
        p = Path(str(data_path).replace('"',''))
        
        if filetype in ['Archive', 'Monument', 'Excel']:
            if 'archive' in p.name.lower():
                folder = self.archive
                self.sortfolder(p, filetype, folder)
            elif 'monument' in p.name.lower():
                folder = self.monument
                self.sortfolder(p, filetype, folder)
            else:
                for subdir in p.rglob('*'):
                    #print('subdir: ' + str(subdir)) #test that this is working
                    if subdir.is_dir():
                        if 'archive' in subdir.name.lower():
                            folder = self.archive
                            self.sortfolder(subdir, filetype, folder)
                        elif 'monument' in subdir.name.lower():
                            folder = self.monument
                            self.sortfolder(subdir, filetype, folder)
                        else:
                            pass   
        else:
            if p.suffix in ['.docx', '.pdf']:
                self.docs.append(str(p))
                self.dirpath = str(p.parent)
            else:
            # Walk subdirectories and files
                if any(subdir.is_dir() for subdir in p.iterdir()):
                    for subdir in p.iterdir():
                        if subdir.is_dir():
                            self.dirpath = str(subdir)
                            for file_path in subdir.rglob('*'):
                                if file_path.is_file():
                                    self.docs.append(str(file_path))
                else:
                    for file_path in p.iterdir():
                        if file_path.is_file():
                            self.dirpath = str(p)
                            self.docs.append(str(file_path))
             
    def runtransl(self, ilanguage, olanguage, data_path, filetype, glossfile, sheet, input_column, start_row, output_dir): #runs the translation methods defined in the translate module
            
        def show_folders(l, type):
            print('length ' + str(len(l)))
            self.l = l
            print('The folder being translated is ' + str(type) + '. The files in this folder are:') #create label widget saying that and listing files
            for f in l: #for each entry in l
                n = int(l.index(f)) + 1
                name = str(str(n) + '. ' + str(f)) #file name is the path + number beforehand
                print(name)  #creates label so files are listed
                            
        
        if len(self.archive) == 0 and len(self.monument) == 0 and len(self.docs) == 0:
            print('Empty input! Did you select the wrong filetype or put in an invalid path? Reset and try again.')

        
        if filetype == 'Excel' or filetype == 'Archive' or filetype == 'Monument': #if the file type is Excel, Archive or Monument
            for l in self.biglist: #for l in biglist
                while True: 
                    if l == self.archive: #if l is archive
                        type = 'Archives'
                        if len(l) == 0:
                            break
                        else:
                            show_folders(l, type)  
                        
                    elif l == self.monument:
                        type = 'Monuments'
                        if len(l) == 0:
                            break
                        else:
                            show_folders(l, type)
                                  

                    else:
                        continue
                        
                        
                      #makes remo global so it can be called in other functions         
                    remy = str(input("\n If you want to remove any spreadsheets, enter the corresponding numbers here separated by comma, otherwise enter N:")) 
                    if remy:
                        self.removefiles(remy,l,data_path,ilanguage,olanguage,filetype, glossfile, sheet, input_column, start_row, output_dir)
                     
                    
        else: #if other types of files - PDF, DOCS
            
            l = self.docs
            self.l = l
            print('The files to be translated are:') #create label widget saying that and listing files
                          
            for f in self.docs: #for each entry in l
                n = int(l.index(f)) + 1
                name = str(str(n) + '. ' + str(f)) #file name is the path + number beforehand
                print (name)
                
    
            remy = input('If you want to remove files, enter corresponding numbers here separated by comma, otherwise enter N:')
            if remy:
                self.removefiles(remy,l,data_path,ilanguage,olanguage,filetype, glossfile,sheet, input_column, start_row, output_dir)
                
    def removefiles(self, remy, l, data_path, ilanguage, olanguage, filetype, glossfile, sheet, input_column, start_row, output_dir):
        '''
        This allows the user to remove any files they want by listing their index - this is useful for files that have already been translated, or that you know have problems, or want to avoid for whatever reason
        The program currently successfully iterates through at least 3 - 4 files. If the program ends up being slow, it might be worth running the shorter-entry files to showcase that it is working
        '''
        #testff = tk.Label(master=framea, text = str(remove) + 'works')
        #testff.pack(fill= tk.X, side = tk.TOP)
        if str(remy) == 'N': #if No files are to be removed
            self.xl_proceed(sheet, start_row, ilanguage, olanguage, input_column, filetype, glossfile, data_path, output_dir)
        elif str(remy) == '': #if it's a blank (clicked OK by accident)
            self.runtransl(ilanguage, olanguage, data_path, filetype,sheet, input_column, start_row, output_dir) #rerun the window
        else: #otherwise
            if ',' in str(remy): #if , is in the string
                rem = remy.split(',') #split values by presence of comma
                i = 1 #i is 0
                for r in rem: #for r in the split string
                    del l[int(r) - int(i)] #delete the file from list using its index calculated by subtracting i from the number given (because the index will update due to deletion i needs to be updated as well)
                    i += 1 #because every time a file is deleted the remaining indices are updated i has to be updated as well
                self.runtransl(ilanguage, olanguage, data_path, filetype, glossfile,sheet, input_column, start_row, output_dir) #runs window with updated files
            elif str(remy).isdigit() == True:
                del l[int(str(remy)) - 1]
                self.runtransl(ilanguage, olanguage, data_path, filetype, glossfile,sheet, input_column, start_row, output_dir)
            else:
                print('Invalid Input')
                self.runtransl(ilanguage, olanguage, data_path, filetype, glossfile,sheet, input_column, start_row, output_dir)    
                

        '''
        
            Here the user is able to put in which sheet the program should access and which columns they would like to translate and in what order. This is left up to the user so that if only a particular 
            column needs to be translated quickly for multiple files, that can be accounted for as well as, having all the needed fields be translated by putting in multiple columns. The format and order are critical here:
            the format because there are columns like AK or BD made up of 2 characters that still need to be read as one entry, but there is also a need to combine data from 2 Russian-entry columns into 1 English-entry
            column. As such if a / is placed then the columns are translated together and output as 1 entry; if a comma is placed, then that means the columns should be translated separately with separate outputs. T
            The order has to be the same for the input and output as that is the only way the program knows which column corresponds to which
            
        '''  
       # table = self.tables[0] #once each file is cycled through the SQL table can be accessed for analysis - in this case the Archive table is used but either can be depending on what
        #is wanted for the analysis
       # SQLite.analysis(SQLite, table, self.SQLcolumns) #call SQLite analysis to find number of different monuments

    def edit_inputs(self, input_sheet, start_row, ilanguage, olanguage, input_column, filetype, glossfile, data_path, output_dir):
        sheet_in = input("Enter 'arg' to keep same input, or enter sheet in spreadsheet (ex: 'Data Sheet'):")
        input_col_in = input('Enter "arg" to keep same input or enter columns to be translated separated by space')
        '''output_col_in = input('Enter "arg" to keep same input or enter columns into which translated data is put')''' 
        start_row_in = input ('Enter "arg" to keep same input or enter row number from which to start translation (ex: 5 - exclude column names):')

        if sheet_in.lower() != 'arg':
            input_sheet = sheet_in

        if input_col_in.lower() != 'arg':
            input_column = input_col_in.split()  # Assuming input as space-separated values

        '''if output_col_in.lower() != 'arg':
            output_column = output_col_in.split() ''' # Assuming input as space-separated values

        if start_row_in.lower() != 'arg':
            start_row = int(start_row_in)

        if input_sheet == 'Monuments':
            sheet = 'Data Sheet'
        elif input_sheet == 'Archive':
            sheet = '2.Описание'   
        else:
            sheet = input_sheet

        self.xl_proceed(sheet, start_row, ilanguage, olanguage, input_column, filetype, glossfile, data_path, output_dir)
    
    def xl_proceed(self, input_sheet, start_row, ilanguage, olanguage, input_column, filetype, glossfile, data_path, output_dir):
        
        transl = TranslateRun()
        
        txt1 = 'Sheet: ' + str(input_sheet) + '\n' +'Columns to translate: ' + '\n'
        for i in input_column:
            txt2 = str(i) +  '\n'
            txt1 = txt1 + txt2
         
        print(txt1 + '\n Starting row: ' + str(start_row)) 
        def runn():
            
            try:     
                for file in self.l:
                    
                    print('Working on file')
                                     
                    if file in self.archive:
                    #setting which row to start on - this can easily also be a user input
                        cols = 'C' #setting which column is used for counter - it has to be a column which will always have data in it regardless of entry so it is set to the CAAL ID column which has 
                    #to be filled for the entry to exist in the spreadsheet (as if it has no CAAL ID it should not be getting recorded in the first place) 
                    else:
                    #ibid - setting it here but can easily be changed through user input
                        cols = 'G' #sets G as the column because it is the CAAL ID
                
                    row_length = pd.read_excel(file, input_sheet, usecols = cols) #creates panda dataframe to figure out how many rows there are in the sheet
                    #column D is used bc it is a pre-filled column that has to be filled for the row to exist
                    #which is then needed to find the max_row so it can be used in ranges

                    max_row = int(len(row_length.dropna())) + 3#defines max_row (to be used in ranges as the length of row_length dataframe - as the length of the dataframe with the extracted column)
                    #dropna is needed because many of the spreadsheets have hundreds of thousands of blank cells loaded in below the data that end up being counted otherwise which makes everything MUCH slower
                    #but because it drops any blank values the column for it has to be one that will have values for every single entry - so the CAAL_ID is chosen an entry has to have a CAAL ID to be entered
                    #in the first place
                    if int(len(row_length)) < int(start_row): #to skip over empty files
                        continue
            
                    else:
                        #print('working')
                        transl.runtransl(file, input_sheet, input_column, int(start_row), int(max_row), ilanguage, olanguage, filetype, glossfile, output_dir)
                        #transl.SQLinput(file, self.folder, input_sheet, self.tables, data_cols, max_row)
                    
                for l in self.biglist:
                    l.clear()
                print('Successful - rerun for next folder')
           
            except PermissionError:
                print('If Permission Error check permissions on files and tick "allowed to edit" for everyone or check that files are closed')
                    
                reset = str(input("Reset? Y/N:"))
                if reset == 'Y':
                    self.runtransl(ilanguage, olanguage, data_path, filetype, glossfile, input_sheet, input_column, start_row)
                else:
                    sys.exit(1)
                    
            except Exception as e:
                print ('ERROR:' + str(e) + ' ' + str(e.args) + '/n' + str(traceback.print_exc(limit=None, file=None, chain=True)))

                reset = str(input("Reset? Y/N:"))
                if reset == 'Y':
                    self.runtransl(ilanguage, olanguage, data_path, filetype, glossfile, input_sheet, input_column, start_row)
                else:
                    sys.exit(1)
                    
                        
                
        proceed = str(input('\n Confirm Y to continue, N to go back:'))
        if proceed == 'Y':
            runn()
        else:
            self.edit_inputs(input_sheet, start_row, ilanguage, olanguage, input_column, filetype, glossfile)
            
        proceed = str(input('\n Confirm Y to continue, N to go back:'))
        if proceed == 'Y':
            runn()
        else:
            self.edit_inputs(input_sheet, start_row, ilanguage, olanguage, input_column, filetype, glossfile, data_path)
    
class Begin: 
    def modrun(self, data_path, ilanguage, olanguage, filetype, glossfile, sheet, input_column, start_row, output_dir):
             #user input of data path
            #language = 'ru' #can change this to Turkmen/Uzbek/Chinese - in future will be dropdown list for user input
            
            load = LoadFolders() #the load folder class
            load.openfolder(data_path, filetype) #opens the folders in the data path
            load.runtransl(ilanguage, olanguage, data_path, filetype, glossfile, sheet, input_column, start_row, output_dir) #runs translation

                     
