'''
Created on 1 Dec 2021

@author: turch
'''
from deep_translator import GoogleTranslator as transl #uses deep_translator instead of googletrans because
# googletrans maxes out times that you can run entries - this has a much higher threshold
#although if it starts timing out lots can change to MicrosoftTranslator - still effective 
#import openpyxl as xl
import xlwings as xw
import pandas as pd
import fnmatch as fn
import os
from pathlib import Path
import traceback

class TranslMethods:
    
     #creates list for untranslated strings
    checked_untrans = [] #creates list for untranslated strings checked against the glossary
    translated = [] #creates list for translated strings
    combineddata = []
    rows = 0
    err_txt = 'Error status: Functioning fine'
    ldat_txt = '...'
    extr_txt = '...'
    chk_txt = '...'
    counter1_txt = '...'
    chk2_txt = '...'
    counter2_txt = '...'
    com_txt = '...'
    ent_txt = '...'
    chk3_txt = '...'
    resett = False
    txt_lst = [err_txt, ldat_txt, extr_txt, chk_txt, counter1_txt, chk2_txt, counter2_txt, com_txt, ent_txt, chk3_txt]
    
        
    def load_data(self, filename, input_sheet, input_column, start_row, max_row):
        modulepath = os.path.dirname(xw.__file__)
        print(modulepath)
        #max_row = 10 #FOR TESTING ONLY - REMEMBER TO UNDO AFTERWARDS (sets max to 5 so can run through code faster)
        col = ','.join(input_column) #joins column names as a string with , - needed to set the columns for the dataframe
        #print('max row' + str(max_row))
        rows = int(int(max_row) - int(start_row)) #number of rows is max_row minus start row - b/c max row refers to all of the filled 
        self.rows = rows
        
        self.ldat_txt = 'The length of data is ' + str(rows) + ' rows. The starting row is ' + str(start_row) +'\n'
        print(self.ldat_txt)
        
        skipline = int(start_row) - 1 #this is done b/c pd starts with 0-index so 0 = 1st line, etc. and we want to skip the lines before the start row (so if starting row is 5 we want to skip 4 lines (0,1,2,3)
        data = pd.read_excel(filename, input_sheet, header=None, names=input_column, usecols=col, skiprows = skipline, nrows = rows) #looked at pandas documentation for how to work with excel
        print("Data read from Excel:")
        print(data.head())
        data = data.fillna('') #fills in the empty cells with empty strings so that they can be processed properly
        
        '''This part of code grabs the column names in the spreadsheet to link them to the column letters provided.
        The reason for this is that later only certain columns need to be concatenated - rather than calling by the column 
        letters provided, as the order of columns can change, they're instead called by the actual names given in the
        spreadsheet. But this means these need to be linked to the data.'''
        names = pd.read_excel(filename, input_sheet, header=1, usecols = col, nrows=1) #creates a separate dataframe for the names of columns (which are not grabbed in the original dataframe due to the starting row being at the data, rather than the columns
        headers = list(names.columns) #sets it so can be referenced outside of the function
        transl_columns = [] 
        self.transl_columns = transl_columns
        for h in headers:
            transl_h = transl(source='ru',target='en').translate(h)
            if 'name' in str(transl_h).lower():
                transl_h = 'Name' #this is done because the 'Name' column can be a very long string, which gets more difficult to call later
            elif 'quantity' in str(transl_h).lower():
                transl_h = 'Quantity'
            elif 'caal' in str(transl_h).lower():
                transl_h = 'CAAL_ID'
            elif 'passport' in str(transl_h).lower(): #some of the spreadsheets have 'passport of history of USSR' as a column name instead of Description
                transl_h = 'Description'
            transl_columns.append(transl_h)
        linked_cols = dict(zip(input_column, transl_columns)) #creates a dictionary with the original column names as keys and the translated column names as values
        self.linked_cols = linked_cols 
        print(linked_cols)

        untrans_to_dict = data[input_column].to_dict(orient = 'records')
        untranslated = [{linked_cols.get(k, k): v for k, v in entry.items()} for entry in untrans_to_dict]
        self.untranslated = untranslated
        
        
        
        self.extr_txt = 'Data extracted'
        print(self.extr_txt)
        #print(self.untranslated)
    
    def create_glossary(self, ilang, glossfile): 
        
        '''
        here we need to create a dictionary which the untranslated data can be compared against using the thesaurus spreadsheet developed by CAAL. The ditionary is created by using pandas to read the spreadsheet
        and make a dataframe of it, converting it to a CSV file as for some reason the Russian here did not want to encode otheriwse, and creating a dictionary with the Russian words set as keys and the English words
        as values; this means that strings can then be compared against the keys to see if they are present, and then substitute the keys for the values.
        
        '''
        
        project_folder = os.path.split(os.path.dirname(os.path.abspath(__file__)))[0] #splits path from the directory name to the main module being opened here
         #joins the main module path with that of data folder so it can be accessed
        #various StackOverflow queries helped with understanding errors that were arising here, particularly in terms of encoding
        if ilang == 'ru':
            
            rus = pd.read_excel(glossfile, usecols = 'B:C')#uses pandas to read excel file extracting columns B to C (Russian and English)
            rus.to_csv("Rusgloss.csv", index = None, header=True) #converts excel to csv because otherwise there are encoding issues that cannot be addressed in pandas read_excel because it removed ability to specify encoding
            rusdf = pd.DataFrame(pd.read_csv("Rusgloss.csv", encoding = 'utf-8')) #reads created csvfile to dataframe specifying the encoding so the Russian characters are read properly
            russgloss = rusdf.set_index('Russian').T.to_dict('list')
            #print(russgloss)
            self.russgloss = russgloss #sets it so it can be referred
            print('Russian glossary created')
            return russgloss
        elif ilang == 'zh_CN':
            cn = pd.read_excel(glossfile, usecols = 'C:D')
            cn.to_csv('Chinese_gloss.csv', index = None, header = True)
            cndf = pd.dataframe(pd.read_csv('Russgloss.csv', encoding = 'utf-8'))
            cngloss = cndf.set_index('Chinese').T.to_dict('list')
            self.cngloss = cngloss
            print('Chinese glossary created')
            return cngloss
        
    def data_check(self, ilang, olang, filetype, glossfile):
        #THIS WHOLE FUNCTION NEEDS REDOING - don't use right now, it's not relevant for Archive spreadsheets anyways
        '''
        This function is to check the untranslated data, both for spelling using spellcheck function and against the dictionary created in create_glossary (thesaurus of CAAL terms) 
        so that terms which appear in the Russian glossary (keys) are replaced by the English values in each string. This is done to avoid inconsistent Google Translate translations, 
        which can sometimes translate one word into different things, particularly specialist terms such as the archaeological ones used here. As such, the glossary was made to allow for consistent tnanslation.
        
        The list of untranslated strings is iterated through, with empty lists for the obtained patterns and replacements. Each key in the glossary 
        is iterated through, with '*' wildcard characters added to the key, and the '[]' characters removed from the replacement string (as a str of the value would return it within brackets)
        Fnmatch is used to match the untranslated string un with the set pattern - I used fnmatch over re because it requires less specification, and it returns a boolean value which allows the check.
        If the value is True, i.e. if the pattern string matches up with the untranslated string, the key and replacement strings are added to a list for replacement. Once the glossary is iterated through,
        the reps and pats lists should contain an equal number of the words/phrases and desired replacements, which are then iterated through to replace within the string.
        
        The checked string (newstring) is then recapitalised according to which column it is in (if it is a Name column, all the individual words are capitalised - if anything else, the string is split up according to sentences and recapitalised)
        and then added to the checked_untranslated list.
        
        '''  
        if str(ilang) == 'ru' and str(olang) == 'en':
            gloss = self.create_glossary(ilang, glossfile)
        elif str(ilang) == 'zh_CN' and str(olang) == 'en':
            gloss = self.create_glossary(ilang, glossfile)
        elif str(olang) == 'ru' and str(ilang) == 'en':
            gloss = self.create_glossary(olang, glossfile)
        elif str(olang) == 'zh_CN' and str(ilang) == 'en':
            gloss = self.create_glossary(olang, glossfile)
        else:
            gloss = {'1':['1']}
        
        self.chk_txt = 'Checking data for spelling errors and proper nouns'
        
        keep_capital = ['основное имя', 'название', 'aвтор', 'аддрес', 'Primary Name', 'Primary Address'] #defines list of columns to keep capitalised - includes 'Main Name', 'Name', 'Author', 'Address' columns which may be in various spreadsheets
        '''
        this is needed to determine if all the words in the string need to be capitalised after data check, which lowercases the whole string. If the string is part of the names column, every word needs to be capitalised
        so that the names of places are properly formatted. If the string is not part of names/addresses i.e. contains multiple sentences, then it will be split up into sentence-strings and each of those will be capitalised
        as I have yet to find a method to restore capital letters as they were.
        '''
        for col in self.names.columns: #for each column in the names dataframe created in input_data
            if any(c in col.lower() for c in keep_capital): #if a part of the lowercase column name matches any string in the keep_capital list
                capitalise = True #then capitalise is set to True
            else: #otherwise capitalise is set to False
                capitalise = False
        
        #print(capitalise) to check that it is registering as True/False - it is
        n=0       
        for data_dict in self.untranslated: #data for each entry is contained in a dictionary, so for each dictionary in the list of untranslated data
            untrans_data ={}
            #print('Dictionary: ' + str(data_dict))
            for col, key in data_dict.items(): #for each key in the dictionary
                un = str(key)
                #print('initial string: ' + str(un))
                if un == '' or un is None:
                    untrans_data[col] = un
                    n=n+1
                    self.counter1_txt = str(n) + ' entries checked'
                    #print(self.counter1_txt)
                elif str(un) == 'nan':
                    untrans_data[col] = un = 'N/A'
                    n=n+1
                    self.counter1_txt = str(n) + ' entries checked'
                    #print(self.counter1_txt)
                else:
                    
                    reps = []
                    pats = [] #creates lists for replacements and patterns that are found through the check against the glossary
                    
                    #print('working on it')
                    if ilang == 'ru' and olang == 'en':
                        for key in gloss: #for each key in the russian glossary dictionary - so for each Russian term
                            pattern = '*' + str(key).lower() + '*' #the pattern is * wildcard character + the key + * - so it matches wherever the key may be present in the string. Fnmatch is used so * is accepted as a wildcard
                            replacement = str(gloss[key]).lower()[2:-2] #the replacement string which contains the English translation of the Russian key - lowercased as the whole string will be
                            if fn.fnmatch(un, str(pattern)) == True: #fnmatch automatically casematches, so lowers un and the pattern - if there is a match, so if boolean = True
                                reps.append(str(replacement)) #then append the replacement value to reps
                                pats.append(str(key).lower()) #and append the Russian word (which has a match) to pats - so the two still correspond with each other by index - lowercased as the rest of the string will be so if it isn't it won't get a match
                                continue
                            else: #otherwise reset the loop and continue
                                continue
                        
                        for i in range(0, int(int(len(pats)) - 1)): #if i = 0 to the length of pats (-1 because len starts from 1 while index starts from 0)
                            un = un.lower().replace(pats[i], reps[i]) #un becomes the lowercased string, with the pattern replaced
                            
                    elif olang == 'ru' and ilang == 'en':
                        
                        for key in gloss:
                            pattern = '*' + str(gloss[key]).lower() + '*'
                            replacement = str(key).lower()[2:-2]
                            if fn.fnmatch(un.lower(), str(pattern)) == True: #fnmatch automatically casematches, so lowers un and the pattern - if there is a match, so if boolean = True
                                reps.append(str(replacement)) #then append the replacement value to reps
                                pats.append(str(gloss[key]).lower()) #and append the Russian word (which has a match) to pats - so the two still correspond with each other by index - lowercased as the rest of the string will be so if it isn't it won't get a match
                                continue
                            else: #otherwise reset the loop and continue
                                continue
                        for i in range(0, int(len(pats) - 1)): #if i = 0 to the length of pats (-1 because len starts from 1 while index starts from 0)
                            un = un.lower().replace(pats[i], reps[i]) #un becomes the lowercased string, with the pattern replaced
                    else:
                        pass
                    
                    nu = un.replace('.,', ',') #removes duplicated punctuation which can appear after the data check
                    newstring = nu.replace('..', '.') #removes duplicated punctuation appearing after data check
        
                    '''
                    The following method is somewhat finicky and could be improved on, but involves the recapitalisation of the string after it was made all-lowercase to match properly with the glossary terms
                    It is finicky because there are multiple sentences within strings, so .capitalize() on the whole string doesn't work because it uppercases the first letter but makes the rest of the characters lowercase - we want each
                    sentence, as well as names, capitalised. So instead the string is split up along '.' which encapsulates both sentences and names when initials and last names are given (as is the standard in these
                    descriptions) and then capitalized, before being joined back together and appended. There is also a very specific method intended to capitalise the initials as well by checking if the last word of 
                    the list of words in a sentence is only 1 character (which should only be with initials) and capitalising that
                    I will continue looking for ways to improve on this but this is the most thorough one I have come up with.
                    '''
                    if capitalise == True: #if capitalize is set to True from earlier
                        caps = newstring.strip().title()#then caps is the newstring with every word capitalised (.title())
                        #print(newstring)
                        untrans_data[key] = caps
                        n += 1
                        self.counter1_txt = str(n) + ' entries checked'
                        #print(self.counter1_txt) #and appended to checked_untrans list for translation
                    else: #otherwise (if capitalise is False)
                        nu = newstring.split('.') #splits newstring along '.' as separators so it's individual sentences
                        cap = [] #creates empty list to store these in
                    
                        for sent in nu: #for each sentence in the nu string
                            caps = sent.strip().capitalize() #strips leading/tailing spaces
                            sep = caps.split(' ') #splits this string along ' ' - so now a consists of all the individual words within sentence z - this is done so initials are capitalised as they end up being lowercase 
                            #because the above-splitter is along '.' so things like 'V.Katsev' will become split
                            last = len(str(sep[-1])) #v is the length of the last word within the list of words in a
                            #if the last word is only one character (which if Russian grammar/spelling rules are followed should only be in the case of initials i.e. 'house of v' and 'Katsev' are separate when it should be house of V.Katsev
                            if last == 1: 
                                capp = sep[-1].capitalize() #then the last word is capitalised
                                sep.pop(-1) #the previous last word is removed
                                sep.append(capp) #and the new capitalised initial is added instead
                                joined = ' '.join(sep) #the string is joined back together using ' ' as separator
                                cap.append(joined) #and appended
                            elif last > 1: #otherwise if the last word is more than one character
                                cap.append(caps) #append the capitalised sentence as it was
                            
                        fin = '. '.join(cap) #rejoins the sentences in cap with a period 
                        
                        #print('1.' + un + '\n 2.' + check_un + '\n 3.' + fin) #check
                        untrans_data[col] = fin
                        print('')
                        n += 1
                        self.counter1_txt = str(n) + ' entries checked'
                        #print(self.counter1_txt) 
            self.checked_untrans.append(untrans_data) #appends to checked untranslated
        print('UNTRANS CHECKED: ' + str(self.checked_untrans))
           
                
    def translator (self, ilanguage, olanguage, filetype):
        if filetype == 'Archive':
            untranslated = self.untranslated #bypasses the data check function so far and uses the original untranslated list of dictionaries
            #this is because the Archive data does not need to be checked against the glossary as much as it has less specialist terms that need to be translated consistently
        else:
            untranslated = self.checked_untrans

        self.chk2_txt = 'Now translating - this can take a while, make a cup of tea :)'
        print(self.chk2_txt)
        try:
            n=0
            for data_dict in untranslated:
                transl_data = {} #creates empty dictionary for the translated data
                for col,item in data_dict.items(): #for strings in untranslated
                    if col.lower() == 'caal_id':
                        transl_data[col] = item #we don't want to translate the CAAL ID as it's already in English
                        #repeat for any other columns that should not be translated
                    else:
                        transl_text = str(item) #converts it into a string (to be sure that it is actually coming out as string value)
                        transltxt = transl(source=ilanguage, target=olanguage).translate(transl_text) #translates the string - taken from deep_translator documentation#
                        transl_data[col] = transltxt
                            #so progress can be checked - and can see if it gets stuck on any particular entries
                        #print(transltxt) #optional - can uncheck to see the translated results
                        #print('translating2')
                        #print(i + '\n \n' + transltxt +'\n \n')
                self.translated.append(transl_data)
                #print("If it's not proceeding, check if Excel has opened the file because it contains macros, click 'x' on the macros to access the spreadsheet and program will resume")
                print("translated: " + str(transl_data)) #check that the translation worked
                n= n+1
                self.counter2_txt = str(n) + ' entries translated'
                print(self.counter2_txt)
        except Exception as e: #except if there is an error 
            self.err_txt = str(traceback.print_exc())
            print(str(traceback.print_exc()))
            
            #print('''IF THE CONNECTION HAS TIMED OUT, WAIT SOME TIME/TRY AGAIN OR TRY PUTTING IN headers={
                                    #'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_9_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/35.0.1916.47 Safari/537.36'}
                                   # INTO deep translator's google_trans translate function in response = requests.get on line 117 ''')
 
        #print("If it's not proceeding, check if Excel has opened the file because it contains macros, click 'x' on the macros to access the spreadsheet and program will resume")    


    def combinedata(self):
        '''
        This function is for columns like Location Notes, where the English translation is input in the same entry underneath the Russian data rather than in its own column. So we need to combine
        the untranslated and translated data into one entry, which can then be input into the column
        '''
        list_combined = []
        self.list_combined = list_combined
        combine = ['Name', 'Description']
        #creates a list of column names that need to be combined - this is done so that the untranslated and translated data can be combined in the same column

        for data_dict in self.translated:
            combined_data = {} #creates empty dictionary for combined data
            for col, item in data_dict.items(): #for each column and item in the translated data
                id = data_dict.get('CAAL_ID')
                untrans = next((d for d in self.untranslated if d.get('CAAL_ID') == id))
                
                if col in combine: #if the column is in the combine list
                    #but then it needs to match column name perfectly
                    #this pulls out the corresponding untranslated dict from the list of untranslated dicts based on matching CAAL_ID
                    original = untrans[col]
                    
                    combined = str(original) + ' / ' + str(item) #combines untranslated and translated data together
                    
                    combined_data[col] = combined #adds the combined string back to dictionary with the same column name
                else: #otherwise just add the translated data as it is
                    combined_data[col] = item #adds the translated data to the dictionary
                    
            final_data = {'Title': str(combined_data['CAAL_ID']) + ' / ' + str(combined_data['Name']),
                          'Description': combined_data['Description'],
                          'Type': str(combined_data['Type of material']) + '. ' + str(combined_data['Quantity'])
                          }
            
            #print(final_data)
            list_combined.append(final_data) #appends the combined dictionaries

    def input_data(self, filename, input_sheet, input_column, start_row, max_row):

        '''''
        This function inputs the translated data from the list into the cells using xlwings as xw - it opens the Sheet, inputs the data using a combination of the input_column and given row, knowing that the data is in
        the same order because that's the order it was extracted in and worked on.
        
        The below code emerged from combined xlwings documentation and various tidbits of code obtained from StackOverflow as I was running into errors.
        '''''
        wb = xw.Book(filename) #opens the workbook with the filename
        sheet = wb.sheets[input_sheet] #opens the sheet with the input_sheet name
        
        i = 0  
        for row in range(int(start_row), int(len(self.translated)) + 1): #for row in the range of starting row to max_row
            cell = str(input_column + str(row)) #the cell is the string combination of input_column and row - need to convert row to string so it can be concatenated with column string
            #allowing you to define which cell the data is going in
            #print(cell)
            if self.combine == False:
                sheet.range(cell).value = self.translated[i] #the xw command for writing into a cell where value is defined as i in translated
            else:
                sheet.range(cell).value = self.combineddata[i]
            i = i + 1 #i adds 1 to iterate through the list  
            self.chk3_txt = str(i) + ' entries input'
            print(self.chk3_txt)
                #this prevents the already-inputted cell values from being overwritten by future iterations
                
        wb.save() #saves the changes
        wb.close() #closes the workbook to allow next columns to be translated
 
    def dict_to_csv(self, filename, output_dir):
        csv_file = str(Path(filename).stem) + '_translated_data.csv' #the name of the csv file to be created
        df = pd.DataFrame(self.list_combined) #creates dataframe of combined data
        print(self.list_combined)
        # Save to CSV
        output_folder = output_dir / 'translated_csv' #creates folder in output_directory for translated csvs
        output_folder.mkdir(parents=True, exist_ok=True)  # Ensure the output directory exists
        csv_path = output_folder / csv_file #creates csv path
        df.to_csv(csv_path, index=False, encoding='utf-8-sig') #makes csv file
        print('Successfully produced CSV')


    def reset(self): #resets the function so there are no leftover values in the lists for the future values
        self.untranslated.clear() #clears the untranslated list
        self.checked_untrans.clear() #ibid for checked untranslated
        self.translated.clear() #ibid for translated
        self.combineddata.clear()
        self.ldat_txt = '...'
        self.extr_txt = '...'
        self.chk_txt = '...'
        self.counter1_txt = '...'
        self.chk2_txt = '...'
        self.counter2_txt = '...'
        self.com_txt = '...'
        self.ent_txt = '...'
        self.chk3_txt = '...'
       
    
class TranslateRun: #the class to run the above methods
    transl = TranslMethods() #transl is the TranslMethods class
        
    def runn(self, filename, input_sheet, input_column, start_row, max_row, ilanguage, olanguage, filetype, glossfile, output_dir): #the whole translation and SQL input process with the associated variables
        self.transl.load_data(filename, input_sheet, input_column, start_row, max_row)
        #self.transl.data_check(ilanguage, olanguage, filetype, glossfile)
        self.transl.translator(ilanguage, olanguage, filetype)
        self.transl.combinedata()
        #self.transl.input_data(filename, input_sheet, input_column, start_row, max_row)
        self.transl.dict_to_csv(filename, output_dir)
    
    def SQLinput(self, filename, folder, input_sheet, tables, data_cols, max_row):
        self.transl.sql_work(filename, folder, input_sheet, tables, data_cols, max_row)
        
    
    def runtransl(self, filename, input_sheet, input_column, start_row, max_row, ilanguage, olanguage, filetype, glossfile, output_dir):

        print('Now translating: ' + str(filename) + '\n Columns: ' + str(input_column))
        print('Note: If file contains forms/macros you will need to manually close/log into these when they pop-up during entry')
        
        
        self.runn(filename, input_sheet, input_column, start_row, max_row, ilanguage, olanguage, filetype, glossfile, output_dir)
        self.transl.reset()
    
        