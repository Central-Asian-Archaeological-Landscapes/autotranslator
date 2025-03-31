'''
Created on 18 Dec 2021

@author: turch
'''


from src.translate_excel import TranslateRun
from src.translate_doc import TranslDocRun
import os
import pandas as pd
from src.spellcheck import Spelling
from src.sqlanalysis import SQLite
from src.resetter import ResetGUI
import tkinter as tk
from tkinter import *
import traceback
import threading


class loadFolders: #creates a class to load folders within the given path for translation of all files within that folder
    archive = [] #creates list for archive spreadsheets
    monument = [] #and list for monument ones
    docs = []
    biglist = [archive, monument] #biglist for the two above lists so they can be cycled through
    l = None
    
    def openfolder(self, data_path, filetype): 
        '''
        This function is the first to be executed in the program as it locates the folders from the specified data path, accesses the files within them, and extracts the file paths to the specified lists (optional ultimately
        but useful for this group of spreadsheets). The file paths can then be used to access individual files for translation.
        '''
        if '"' in str(data_path):
            data_path = str(data_path).replace('"', '')
            print('w:)')
            
        else:
            print('w:(')
            pass
            
        if filetype == 'Excel':
                if 'archive' in str(data_path).lower():
                    #print('i')
                    if '.xl' in str(data_path):
                        self.archive.append(str(data_path))
                    for data_path, subdirs, files in os.walk(data_path):
                        for f in files:
                            file_path = os.path.join(data_path, f)
                            self.archive.append(str(file_path))
                        for s in subdirs:
                            subdir_path = os.path.join(data_path, s)
                            print(subdir_path)
                            for f in files: #for the directory, subdirectories, and files in the directory path
                                file_path = os.path.join(subdir_path, f) #join file name with the directory path as file path
                                print(file_path)
                                self.archive.append(str(file_path)) #and append that to the archive list 
            
                elif 'monument' in str(data_path).lower():
                    #print('x')
                    if '.xl' in str(data_path):
                        self.monument.append(str(data_path))
                    else:
                        for data_path, subdirs, files in os.walk(data_path):
                            for f in files:
                                file_path = os.path.join(data_path, f)
                                self.monument.append(str(file_path))
                            for s in subdirs:
                                subdir_path = os.path.join(data_path, s)
                                print(subdir_path)
                                for f in files: #for the directory, subdirectories, and files in the directory path
                                    file_path = os.path.join(subdir_path, f) #join file name with the directory path as file path
                                    print(file_path)
                                    self.monument.append(str(file_path)) #and append that to the archive list 
                        
                        
                else:
                        for data_path, subdirs, files in os.walk(data_path):
                            for s in subdirs: #for each subdirectory
                                dirpath = os.path.join(data_path, s) #creates a directory path to add the file name to for a complete file path
                                print(dirpath)
                                if 'archive' in str(dirpath).lower(): #if 'ARCHIVE' is in the subdirectory name
                                    for dirpath, subdirs, files in os.walk(dirpath): #for the directory, subdirectories, and files in the directory path
                                        for f in files:
                                            file_path = os.path.join(dirpath, f)
                                            if str(file_path) in self.archive:
                                                pass
                                            else:
                                                self.archive.append(file_path)
                                            
                                        '''
                                        for s in subdirs:
                                           subdir_path = os.path.join(dirpath, s)
                                           for file in files: #for each file
                                               file_path = os.path.join(subdir_path,file) #join file name with the directory path as file path
                                               print(file_path)
                                               self.archive.append(str(file_path)) #and append that to the archive list
                                               '''
                                            
                                elif 'monument' in str(dirpath).lower(): #otherwise if 'MONUMENT' is in the subdirectory path:
                                    for dirpath, subdirs, files in os.walk(dirpath): #do the exact same process but append it to the monument list instead
                                        for f in files:
                                            file_path = os.path.join(dirpath, f)
                                            if str(file_path) in self.monument:
                                                pass
                                            else:
                                                self.monument.append(file_path)
                                            
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
                
          
     
    def makesql(self, data_path): #makes an SQL from sqlanalysis module
        '''
        This is the first step to loading all the data into an SQL for further analysis. Because we are iterating through files, but we want to compare across files, the SQL database is made now
        so that the tables already exist for the folders, but will be filled with data for each file in the translate methods
        '''
        folder = 'CAAL SQL' #the database name
        self.folder = folder
        
        ascii_conv = [] #creates list for ascii converted values
        for c in str(data_path): #for character in data_path
            ascii_conv.append(ord(c)) #find its ascII value and append it to the table
        uniq_table = sum(ascii_conv) #uniq_table value is the sum of all the ascII
        
        
        tables = [str('Archive' + str(uniq_table)), str('Monument' + str(uniq_table))] 
        
        '''the tables for archive and monument data - the uniq_table value is added to make each table name unique so that
        they don't have to be deleted across different running of the program unless they are the same. I wanted to find a value that would be unique
        for each data path across different times the program is run without having to use the data_path itself as that is quite long. So by getting all the  
        ASCII values and adding them this will be a unique value because the data paths will vary in at least one character. This will allow to save the
        data obtained from the translated spreadsheets within an SQL database, with the files within one folder stored within one table.
        Then when more files have been translated the tables can all be united in a separate program and extensive data analysis can be run
        For now the data analysis is very rudimentary as the file translation is still being optimised'''
        
        self.tables = tables
        
        SQLcolumns = 'Row TEXT, Name TEXT, Description TEXT, Location TEXT' #name, description, and location files - will be used to run some analysis of how many entries come from where, what type they are (based on common words)
        self.SQLcolumns = SQLcolumns
        
        lite = SQLite(folder) #lite is the class SQLite in sqlanalysis - runs the init module with CAAL SQL as the database
        for table in tables: #for each string in tables
            lite.makesql(table, SQLcolumns) #make a table with that as the name, and the given columns
        
                
    def runtransl(self, ilanguage, olanguage, data_path, window, filetype): #runs the translation methods defined in the translate module
        spell = Spelling() #spell is the Spelling() class from spellcheck
        spell.ru_dict() #runs the making russian dictionary module from spelling()

        
        
        framea = tk.Frame(master=window, width = 300, height=300)
        framea.pack(fill = tk.X, side = tk.TOP)
        
        def alterfile():
            global remove #so remove can be called elsewhere
            remove = remo.get()
            framea.destroy()
            frameb.destroy()
            self.removefiles(remove, l, data_path, ilanguage, olanguage, window, filetype) #calls removefiles function
        
        def resett():
            reset = ResetGUI()
            window.destroy()
            reset.restart()
        
        if len(self.archive) == 0 and len(self.monument) == 0 and len(self.docs) == 0:
            empty = tk.Label(master=framea, text = 'Empty input! Did you select the wrong filetype or put in an invalid path? Reset and try again.')
            empty.pack(fill = tk.X, side = tk.TOP)
            
            reset_btn = tk.Button(master=framea, text = 'RESET', command = resett)
            reset_btn.pack(fill = tk.X, side = tk.BOTTOM)
        
        if filetype == 'Excel':
            for l in self.biglist: #for l in biglist
                while True: 
                    if l == self.archive: #if l is archive
                        print('length ' + str(len(l)))
                        if len(l) == 0:
                            break
                        else:
                            self.l = l
                            
                            arc = tk.Label(master=framea, text = 'The folder being translated is Archives. The files in this folder are:') #create label widget saying that and listing files
                            arc.pack(fill = tk.X, side = tk.TOP)
                            
                            
                            for f in l: #for each entry in l
                                n = int(l.index(f)) + 1
                                name = str(str(n) + '. ' + str(f)) #file name is the path + number beforehand
                                lst = tk.Label(master=framea, text = name) #creates label so files are listed in window
                                lst.pack(fill = tk.X, side = tk.TOP)
                                
                            data_cols = 'F, I, N' #data columns for SQL extraction - corresponding to English name, English description, location notes
                        
                        
                        
                    elif l == self.monument:
                        print('length ' + str(len(l)))
                        if len(l) == 0:
                            break
                        else:
                            self.l = l
                            arc = tk.Label(master=framea, text = 'The folder being translated is Monuments. The files in this folder are:')
                            arc.pack(fill = tk.X, side = tk.TOP)
                            
                            for f in l: #same as previous
                                n = int(l.index(f)) + 1
                                name = str(str(n) + '. ' + str(f))
                                lst = tk.Label(master=framea, text = name)
                                lst.pack(fill = tk.X, side = tk.TOP)
                                
                            data_cols = 'B, AD, AK' #data columns for SQL extraction - corresponding to English name, English description, and location notes
                        
                    else:
                        continue
                        
                        
                              
                    #remove = input("\n If you want to remove any spreadsheets, enter the corresponding numbers here separated by comma, otherwise enter N:") 
                    frameb = tk.Frame(master=window, width = 300, height=150)
                    frameb.pack(fill = tk.X, side = tk.TOP)
                    remy = tk.Label(master=frameb, text = 'If you want to remove any spreadsheets, enter corresponding numbers here separated by comma, otherwise enter N')
                    remy.pack(fill = tk.X, side = tk.TOP) #creates label for entry to begin file removal process if neede
                    global remo #makes remo global so it can be called in other functions
                    remo = tk.Entry(master=frameb) #remo is the entry
                    remo.pack(fill = tk.X, side = tk.TOP) #packs it
                    
                    ok_btn = tk.Button(master=frameb, text = 'OK', command=alterfile)
                    ok_btn.pack(fill = tk.X, side = tk.BOTTOM)
                    window.mainloop()
                    
        else: #if other types of files - PDF, DOCS
            l = self.docs
            self.l = l
            arc = tk.Label(master=framea, text = 'The files to be translated are:') #create label widget saying that and listing files
            arc.pack(fill = tk.X, side = tk.TOP)
                          
            for f in self.docs: #for each entry in l
                n = int(l.index(f)) + 1
                name = str(str(n) + '. ' + str(f)) #file name is the path + number beforehand
                lst = tk.Label(master=framea, text = name) #creates label so files are listed in window
                lst.pack(fill = tk.X, side = tk.TOP)
                
            frameb = tk.Frame(master=window, width = 300, height=150)
            frameb.pack(fill = tk.X, side = tk.TOP)
            remy = tk.Label(master=frameb, text = 'If you want to remove files, enter corresponding numbers here separated by comma, otherwise enter N')
            remy.pack(fill = tk.X, side = tk.TOP) #creates label for entry to begin file removal process if neede
            remo = tk.Entry(master=frameb) #remo is the entry
            remo.pack(fill = tk.X, side = tk.TOP) #packs it
            
            ok_btn = tk.Button(master=frameb, text = 'OK', command=alterfile)
            ok_btn.pack(fill = tk.X, side = tk.BOTTOM)
            window.mainloop()
                
    def removefiles(self, remove, l, data_path, ilanguage, olanguage, window, filetype):
        '''
        This allows the user to remove any files they want by listing their index - this is useful for files that have already been translated, or that you know have problems, or want to avoid for whatever reason
        The program currently successfully iterates through at least 3 - 4 files. If the program ends up being slow, it might be worth running the shorter-entry files to showcase that it is working
        '''
        #testff = tk.Label(master=framea, text = str(remove) + 'works')
        #testff.pack(fill= tk.X, side = tk.TOP)
        if str(remove) == 'N': #if No files are to be removed
            if filetype == 'Excel':
                self.input_window(ilanguage, olanguage, window, filetype) #move to input_window
            else:
                self.output_location(filetype, ilanguage,olanguage,window)
        elif str(remove) == '': #if it's a blank (clicked OK by accident)
            self.runtransl(ilanguage, olanguage, data_path, window, filetype) #rerun the window
        else: #otherwise
            if ',' in str(remove): #if , is in the string
                rem = remove.split(',') #split values by presence of comma
                i = 1 #i is 0
                for r in rem: #for r in the split string
                    del l[int(r) - int(i)] #delete the file from list using its index calculated by subtracting i from the number given (because the index will update due to deletion i needs to be updated as well)
                    i += 1 #because every time a file is deleted the remaining indices are updated i has to be updated as well
                self.runtransl(ilanguage, olanguage, data_path, window, filetype) #runs window with updated files
            elif str(remove).isdigit() == True:
                del l[int(str(remove)) - 1]
                self.runtransl(ilanguage, olanguage, data_path, window, filetype)
            else:
                windowc = tk.Tk()
                inv = tk.Label(windowc, text = 'Invalid Input')
                inv.pack(fill = tk.X, side = tk.BOTTOM)
                self.runtransl(ilanguage, olanguage, data_path, window, filetype)
            
    
    def input_window(self, ilanguage, olanguage, windowa, filetype):
        framea = tk.Frame(master=windowa, width = 150, height=150, padx=5, pady=5)
        framea.pack(fill = tk.X, side=tk.TOP)
        
        sheet = tk.Label(master=framea, text = 'Please enter sheet:')
        sheet_input = tk.Entry(master=framea, width = 75)
        sheet.pack(fill = tk.X, side=tk.LEFT)
        sheet_input.pack(fill = tk.X, side = tk.RIGHT)
        
        frameb = tk.Frame(master=windowa, width = 150, height=150, padx=5, pady=5)
        frameb.pack(fill = tk.X, side=tk.TOP)
        
        col1 = tk.Label(master=frameb, text = 'Please enter columns to be translated with slash if you want them to be translated as one, and with a comma if you want them to be translated separately \n (ex: H/J will translate columns H and J and combine them; AK, A will translate columns AK and A separately)')
        col1_input = tk.Entry(master=frameb, width = 75)
        col1.pack(fill = tk.X, side=tk.LEFT)
        col1_input.pack(fill = tk.X, side = tk.RIGHT)
        
        framec = tk.Frame(master=windowa, width = 150, height=150, padx=5, pady=5)
        framec.pack(fill = tk.X, side=tk.TOP)
        
        col2 = tk.Label(master=framec, text = 'Enter input columns in same order/format that you did columns (if you want data from H/J to be input into I, put I first): \n Note: if same column is put in for input/output, it will enter the translated data into the same cell as untranslated and will keep both')
        col2_input = tk.Entry(master=framec, width = 75)
        col2.pack(fill = tk.X, side=tk.LEFT)
        col2_input.pack(fill = tk.X, side = tk.RIGHT)
        
        framed = tk.Frame(master=windowa, width = 150, height=150, padx=5, pady=5)
        framed.pack(fill = tk.X, side=tk.TOP)
        
        col3 = tk.Label(master=framed, text = 'Enter row number from which to start translation (ex: 5 - exclude column names):')
        col3_input = tk.Entry(master=framed, width = 75)
        col3.pack(fill = tk.X, side=tk.LEFT)
        col3_input.pack(fill = tk.X, side = tk.RIGHT)
        
        framee=tk.Frame(master=windowa, width=150, height=150, padx=5, pady=5)
        framee.pack(fill=tk.X, side = tk.TOP)
        
        frame_list = [framea,frameb,framec,framed,framee]
        
        def get_inputs():
            input_sheet = sheet_input.get()
            input_columns = col1_input.get()
            output_columns = col2_input.get()
            start_row = col3_input.get()
            for f in frame_list:
                f.forget()
            self.runtransl2(windowa, ilanguage, olanguage, input_sheet, input_columns, output_columns, start_row, filetype)
            
        
        ok_btn = tk.Button(master=framee,text = 'OK',command=threading.Thread(target=get_inputs).start)
        ok_btn.pack(fill = tk.X, side = tk.BOTTOM)
            
    def output_location(self, filetype, ilanguage, olanguage, window):
        framea = tk.Frame(master=window, width = 150, height=150, padx=5, pady=5)
        framea.pack(fill = tk.X, side=tk.TOP)
        
        def ndd(selection):
            global newdoc
            newdoc = nd.get()
        
        new_doc = tk.Label(master=framea, text = 'Please enter if translation is in new document or within same one (Word only): \n Note: new documents are output as same filename with added "_translated" and new language')
        new_doc.pack(fill = tk.X, side = tk.LEFT)
        nd = StringVar(framea)
        nd.set("-")
        ndmenu = OptionMenu(framea, nd, 'New document', 'Write in same one', command = ndd)
        ndmenu.pack(fill = tk.X, side = tk.RIGHT)
        
        frameb = tk.Frame(master=window, width = 150, height=150, padx=5, pady=5)
        frameb.pack(fill = tk.X, side=tk.TOP)
        framec = tk.Frame(master=window, width = 150, height = 50)
        framec.pack(fill = tk.X, side = tk.TOP) 
        
        def locationn():
            location = loc.get()
            print(location)
            window.destroy()
            if len(str(location)) == 0:
                self.doc_proceed(filetype, ilanguage, olanguage, newdoc, self.dirpath)
            else:
                self.doc_proceed(filetype, ilanguage, olanguage, newdoc, location)
            
        locate = tk.Label(master=frameb, text = 'Enter folder to which new file is saved, leave blank if same as original file:')
        locate.pack(fill = tk.X, side = tk.LEFT)
        loc = tk.Entry(master=frameb, width = 75)
        loc.pack(fill = tk.X, side = tk.RIGHT)
           
        ok_btn = tk.Button(master=framec, text = 'OK', command=locationn)
        ok_btn.pack(fill = tk.X, side = tk.BOTTOM)
                
    def runtransl2(self, windowa, ilanguage, olanguage, input_sheet, input_columns, output_columns, start_row, filetype):
        
        columns = input_columns
        col_list = columns.split(',') #creates list of columns with , as the splitter

        input_col = output_columns
        inputs = input_col.split(',') #separates input columns based on ,
    
        columndict = dict(zip(col_list,inputs)) #zips the list of columns and inputs together so the 1st column/group of columns becomes associated with the 1st input (as key/value) and so on
        
        self.xl_proceed(windowa, columndict, input_sheet, start_row, ilanguage, olanguage, col_list, inputs, filetype)
        
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
    
    def xl_proceed(self, windowa, column_dict, input_sheet, start_row, ilanguage, olanguage, col_list, inputs, filetype):
        transl = TranslateRun()
        
        windowb = tk.Tk()
        framea = tk.Frame(master=windowb, width = 150, height=150, padx=5, pady=5)
        framea.pack(fill = tk.X, side=tk.TOP)
        
        txt1 = 'Sheet: ' + str(input_sheet) + '\n'
        
        for col in col_list:
            c = int(col_list.index(col))
            txt2 = 'Column to translate: ' + str(col) + '   Column for input: ' + str(inputs[c]) + '\n'
            txt1 = txt1 + txt2
         
        txt = txt1 + '\n Starting row: ' + str(start_row) + '\n Confirm CONTINUE or go BACK:'
        confrm = tk.Label(master=framea, text = str(txt))
        confrm.pack(fill = tk.X, side = tk.TOP)
        
        frameb = tk.Frame(master=windowb, width =150, height=75, padx=5, pady=5)
        frameb.pack(fill = tk.X, side=tk.TOP)
        
        def go_back():
            windowb.destroy()
            self.input_window(ilanguage, olanguage, windowa, filetype)
            
        def runn():
            windowb.destroy()
            
            try:     
                for file in self.l:
                    
                    print('worrk')
                                     
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
                    
                    for key in column_dict: #for each key in the column dictionary (where the data columns and input columns are
                        column_names = []
                        input_column = str(column_dict[key])
                        if '/' in key:
                            column_names = key.split('/')
                            
                
                        else:
                                column_names.append(key)
                        
                        #print('working')
                        transl.runtransl(file, input_sheet, column_names, input_column, int(start_row), int(max_row), ilanguage, olanguage, windowa, filetype)
                    #transl.SQLinput(file, self.folder, input_sheet, self.tables, data_cols, max_row)
                reset = ResetGUI()
                for l in self.biglist:
                    l.clear()
                reset.restart() 
           
                    
            except Exception as e:
                    reset = ResetGUI()
                    windoww = tk.Tk()
                    fram = tk.Frame(master=windoww, width = 150, height = 150, padx = 5, pady=5)
                    fram.pack(fill = tk.X, side = tk.TOP)
                    err = tk.Label(master=fram, text = 'ERROR:' + str(e) + ' ' + str(e.args) + '/n' + str(traceback.print_exc(limit=None, file=None, chain=True)))
                    err.pack(fill=tk.X, side = tk.TOP)
                    tip = tk.Label(master=fram, text = 'If Permission Error check permissions on files and tick "allowed to edit" for everyone or check that files are not open')
                    tip.pack(fill=tk.X, side = tk.TOP)
                    
                    def resett():
                        windowa.destroy()
                        reset.restart()
                        
                    framee = tk.Frame(master=windoww, width = 150, height = 150, padx = 5, pady=5)
                    framee.pack()
                    new_btn = tk.Button(master=framee, text = 'RESTART', command = resett)
                    new_btn.pack(fill = tk.X, side = tk.BOTTOM)
                    windoww.mainloop()
            
        
        back = tk.Button(master=frameb, text = 'BACK', command=go_back)
        back.pack(fill = tk.Y, side = tk.LEFT)
        cont = tk.Button(master =frameb, text = 'CONTINUE', command=runn)
        cont.pack(fill = tk.Y, side = tk.RIGHT)
        windowb.mainloop()
    
    def doc_proceed(self, filetype, ilanguage, olanguage, newdoc, location):
        transl = TranslDocRun()
        reset = ResetGUI()
        windowa = tk.Tk()
        try:
            for file in self.docs:
                transl.run(file, filetype, location, ilanguage, olanguage, newdoc, windowa)
            
            self.docs.clear()  
            reset.restart() 
            windowa.mainloop()
            
        except Exception as e: 
            windoww = tk.Tk()
            fram = tk.Frame(master=windoww, width = 150, height = 150, padx = 5, pady=5)
            fram.pack(fill = tk.X, side = tk.TOP)
            err = tk.Label(master=fram, text = 'ERROR:' + str(e) + ' ' + str(e.args) + '/n' + str(traceback.print_exc(limit=None, file=None, chain=True)))
            err.pack(fill=tk.X, side = tk.TOP)
            tip = tk.Label(master=fram, text = 'If Permission Error check permissions on files and tick "allowed to edit" for everyone or check that files are not open \n If not translating or "connection error" wait some time and try again as this is due to API')
            tip.pack(fill=tk.X, side = tk.TOP)
            
            def resett():
                windowa.destroy()
                reset.restart()
                
            framee = tk.Frame(master=windoww, width = 150, height = 150, padx = 5, pady=5)
            framee.pack()
            new_btn = tk.Button(master=framee, text = 'RESTART', command = resett)
            new_btn.pack(fill = tk.X, side = tk.BOTTOM)
            windoww.mainloop()
    
        
    
class run: 
    def run(self, data_path, ilanguage, olanguage, filetype):
            window = tk.Tk()
            #data_path = data #str(input("Enter full main folder path here: ")) - this is normally a user input but I've made it a preset path for assignment purposes
            #language = 'ru' #can change this to Turkmen/Uzbek/Chinese - in future will be dropdown list for user input
            
            load = loadFolders() #the load folder class
            load.openfolder(data_path, filetype) #opens the folders in the data path
            load.makesql(data_path)
            load.runtransl(ilanguage, olanguage, data_path, window, filetype) #runs translation
            window.mainloop()
                     
