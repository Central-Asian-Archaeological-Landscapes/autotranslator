'''
Created on 18 Dec 2021

@author: turch
'''


from translate_excel import TranslateRun
import os
import pandas as pd
import traceback
import threading


class LoadFolders: #creates a class to load folders within the given path for translation of all files within that folder
    archive = [] #creates list for archive spreadsheets
    monument = [] #and list for monument ones
    docs = []
    biglist = [archive, monument] #biglist for the two above lists so they can be cycled through
    l = None
    
    def sortfolder(self, data_path, filetype, folder):
        if '.xl' in str(data_path):
            folder.append(str(data_path))
        else:
            for data_path, subdirs, files in os.walk(data_path):
                for f in files:
                    file_path = os.path.join(data_path, f)
                    folder.append(str(file_path))
                for s in subdirs:
                    subdir_path = os.path.join(data_path, s)
                    print(subdir_path)
                    for f in files: #for the directory, subdirectories, and files in the directory path
                        file_path = os.path.join(subdir_path, f) #join file name with the directory path as file path
                        print(file_path)
                        folder.append(str(file_path)) #and append that to the archive list 

    def openfolder(self, data_path, filetype): 
        '''
        This function is the first to be executed in the program as it locates the folders from the specified data path, accesses the files within them, and extracts the file paths to the specified lists (optional ultimately
        but useful for this group of spreadsheets). The file paths can then be used to access individual files for translation.
        '''
        if '"' in str(data_path):
            data_path = str(data_path).replace('"', '')
            
        else:
            pass
            
        if filetype == 'Excel':
                if 'archive' in str(data_path).lower():
                    folder = self.archive
                    self.sortfolder(data_path,filetype,folder)
                    
                elif 'monument' in str(data_path).lower():
                    folder = self.monument
                    sortfolder(data_path, filetype, folder)
                        
                else:
                    for data_path, subdirs, files in os.walk(data_path):
                        for s in subdirs: #for each subdirectory
                            dirpath = os.path.join(data_path, s) #creates a directory path to add the file name to for a complete file path
                            print(dirpath)
                            if 'archive' in str(dirpath).lower(): #if 'ARCHIVE' is in the subdirectory name
                                folder = self.archive
                                self.sortfolder(dirpath, filetype, folder)
                                        
                            elif 'monument' in str(dirpath).lower(): #otherwise if 'MONUMENT' is in the subdirectory path:
                                folder = self.monument
                                self.sortfolder(dirpath, filetype, folder)
                                
                            else: #otherwise continue
                                pass
        else:
            if '.docx' in str(data_path) or '.pdf' in str(data_path):
                self.docs.append(data_path)
                dirpath = os.path.dirname(data_path)
                self.dirpath = dirpath
            else:
                for data_path, subdirs, files in os.walk(data_path):
                    if len(list(subdirs)) >0:
                        for s in subdirs:
                            dirpath = os.path.join(data_path, s)
                            self.dirpath = dirpath
                            for dirpath, files in os.walk(dirpath):
                                for file in files:
                                    file_path = os.path.join(dirpath, file)
                                    self.docs.append(str(file_path))
                    else:
                        for file in files:
                            self.dirpath = data_path
                            file_path = os.path.join(data_path, file)
                            self.docs.append(str(file_path))
                    
             
    def runtransl(self, ilanguage, olanguage, data_path, filetype, glossfile): #runs the translation methods defined in the translate module

        def alterfile():
            global remove #so remove can be called elsewhere
            remove = remo.get()
            framea.destroy()
            frameb.destroy()
            self.removefiles(remove, l, data_path, ilanguage, olanguage, filetype) #calls removefiles function
        
        def resett():
            reset = ResetGUI()
            window.destroy()
            reset.restart()
            
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

        
        if filetype == 'Excel':
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
                        
                        
                    global remove  #makes remo global so it can be called in other functions         
                    remove = str(input("\n If you want to remove any spreadsheets, enter the corresponding numbers here separated by comma, otherwise enter N:")) 
                    if remove:
                        self.removefiles(remove,l,data_path,ilanguage,olanguage,filetype, glossfile)
                     
                    
        else: #if other types of files - PDF, DOCS
            
            l = self.docs
            self.l = l
            print('The files to be translated are:') #create label widget saying that and listing files
                          
            for f in self.docs: #for each entry in l
                n = int(l.index(f)) + 1
                name = str(str(n) + '. ' + str(f)) #file name is the path + number beforehand
                print (name)
                
            frameb = tk.Frame(master=window, width = 300, height=150)
            frameb.pack(fill = tk.X, side = tk.TOP)
            global remy
            remy = input('If you want to remove files, enter corresponding numbers here separated by comma, otherwise enter N:')
            if remy:
                self.removefiles(remy,l,data_path,ilanguage,olanguage,filetype, glossfile)
                
    def removefiles(self, remove, l, data_path, ilanguage, olanguage, filetype, glossfile):
        '''
        This allows the user to remove any files they want by listing their index - this is useful for files that have already been translated, or that you know have problems, or want to avoid for whatever reason
        The program currently successfully iterates through at least 3 - 4 files. If the program ends up being slow, it might be worth running the shorter-entry files to showcase that it is working
        '''
        #testff = tk.Label(master=framea, text = str(remove) + 'works')
        #testff.pack(fill= tk.X, side = tk.TOP)
        if str(remove) == 'N': #if No files are to be removed
            self.inputs(ilanguage, olanguage, filetype, glossfile) #move to input_window
        elif str(remove) == '': #if it's a blank (clicked OK by accident)
            self.runtransl(ilanguage, olanguage, data_path, filetype) #rerun the window
        else: #otherwise
            if ',' in str(remove): #if , is in the string
                rem = remove.split(',') #split values by presence of comma
                i = 1 #i is 0
                for r in rem: #for r in the split string
                    del l[int(r) - int(i)] #delete the file from list using its index calculated by subtracting i from the number given (because the index will update due to deletion i needs to be updated as well)
                    i += 1 #because every time a file is deleted the remaining indices are updated i has to be updated as well
                self.runtransl(ilanguage, olanguage, data_path, filetype, glossfile) #runs window with updated files
            elif str(remove).isdigit() == True:
                del l[int(str(remove)) - 1]
                self.runtransl(ilanguage, olanguage, data_path, filetype, glossfile)
            else:
                print('Invalid Input')
                self.runtransl(ilanguage, olanguage, data_path, filetype, glossfile)
            
    def inputs (self,ilanguage,olanguage,filetype, glossfile):
        sheet = input("Please enter sheet in spreadsheet (ex: 'Data Sheet'):")
        input_columns = input('Please enter columns to be translated with slash if you want them to be translated as one, and with a comma if you want them to be translated separately \n (ex: H/J will translate columns H and J and combine them; AK, A will translate columns AK and A separately)')
        output_columns = input('Enter input columns in same order/format that you did columns (if you want data from H/J to be input into I, put I first): \n Note: if same column is put in for input/output, it will enter the translated data into the same cell as untranslated and will keep both')
        start_row = input ('Enter row number from which to start translation (ex: 5 - exclude column names):')
        self.runtransl2(ilanguage, olanguage, sheet, input_columns, output_columns, start_row, filetype, glossfile)
    
                
    def runtransl2(self, ilanguage, olanguage, input_sheet, input_columns, output_columns, start_row, filetype, glossfile):
        
        columns = input_columns
        col_list = columns.split(',') #creates list of columns with , as the splitter

        input_col = output_columns
        inputs = input_col.split(',') #separates input columns based on ,
    
        columndict = dict(zip(col_list,inputs)) #zips the list of columns and inputs together so the 1st column/group of columns becomes associated with the 1st input (as key/value) and so on
        
        self.xl_proceed(columndict, input_sheet, start_row, ilanguage, olanguage, col_list, inputs, filetype, glossfile)
        
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
    
    def xl_proceed(self, column_dict, input_sheet, start_row, ilanguage, olanguage, col_list, inputs, filetype, glossfile):
        transl = TranslateRun()
        
        txt1 = 'Sheet: ' + str(input_sheet) + '\n'
        
        for col in col_list:
            c = int(col_list.index(col))
            txt2 = 'Column to translate: ' + str(col) + '   Column for input: ' + str(inputs[c]) + '\n'
            txt1 = txt1 + txt2
         
        print(txt1 + '\n Starting row: ' + str(start_row)) 
        def runn():
            
            try:     
                for file in self.l:
                    
                    print('Working on file')
                                     
                    if file in self.archive:
                    #setting which row to start on - this can easily also be a user input
                        cols = 'D' #setting which column is used for counter - it has to be a column which will always have data in it regardless of entry so it is set to the CAAL ID column which has 
                    #to be filled for the entry to exist in the spreadsheet (as if it has no CAAL ID it should not be getting recorded in the first place) 
                    else:
                    #ibid - setting it here but can easily be changed through user input
                        cols = 'G' #sets G as the column because it is the CAAL ID
                
                    row_length = pd.read_excel(file, input_sheet, usecols = cols) #creates panda dataframe to figure out how many rows there are in the sheet
                    #column D is used bc for both types of spreadsheets it is a pre-filled column that has to be filled for the row to exist
                    #which is then needed to find the max_row so it can be used in ranges
                    max_row = int(len(row_length.dropna())) + 3#defines max_row (to be used in ranges as the length of row_length dataframe - as the length of the dataframe with the extracted column)
                    #dropna is needed because many of the spreadsheets have hundreds of thousands of blank cells loaded in below the data that end up being counted otherwise which makes everything MUCH slower
                    #but because it drops any blank values the column for it has to be one that will have values for every single entry - so the CAAL_ID is chosen an entry has to have a CAAL ID to be entered
                    #in the first place
                    if int(len(row_length)) < int(start_row): #to skip over empty files
                        continue
            
                    else:
                        for key in column_dict: #for each key in the column dictionary (where the data columns and input columns are
                            column_names = []
                            input_column = str(column_dict[key])
                            if '/' in key:
                                column_names = key.split('/')
                                
                    
                            else:
                                    column_names.append(key)
                            
                            #print('working')
                            transl.runtransl(file, input_sheet, column_names, input_column, int(start_row), int(max_row), ilanguage, olanguage, filetype, glossfile)
                        #transl.SQLinput(file, self.folder, input_sheet, self.tables, data_cols, max_row)
                    
                for l in self.biglist:
                    l.clear()
                print('Successful - rerun for next folder')
           
                    
            except Exception as e:
                    print ('ERROR:' + str(e) + ' ' + str(e.args) + '/n' + str(traceback.print_exc(limit=None, file=None, chain=True)))
                    print('If Permission Error check permissions on files and tick "allowed to edit" for everyone or check that files are closed')
            
                    reset = str(input("Reset? Y/N:"))
                    if reset == 'Y':
                        self.runtransl(ilanguage, olanguage, self.data_path, filetype)
                    else:
                        sys.exit(1)
                        
        def go_back():
            self.inputs(ilanguage, olanguage, filetype)
            
        proceed = str(input('\n Confirm Y to continue, N to go back:'))
        if proceed == 'Y':
            runn()
        else:
            go_back()
    
class Begin: 
    def modrun(self, data_path, ilanguage, olanguage, filetype, glossfile):
             #user input of data path
            #language = 'ru' #can change this to Turkmen/Uzbek/Chinese - in future will be dropdown list for user input
            
            load = LoadFolders() #the load folder class
            load.openfolder(data_path, filetype) #opens the folders in the data path
            load.runtransl(ilanguage, olanguage, data_path, filetype, glossfile) #runs translation

                     
