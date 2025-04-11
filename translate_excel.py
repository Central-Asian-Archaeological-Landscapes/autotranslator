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
import csv
import traceback

class TranslMethods:
    
    #untranslated = [] #creates list for untranslated strings
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
    
    def update(self, err, ldat, extr, chk, counter1, chk2, counter2, com, ent, chk3): #function to allow window to continue updating
        err.config(text = self.err_txt)
        ldat.config(text = self.ldat_txt)
        extr.config(text = self.extr_txt)
        chk.config(text = self.chk_txt)
        counter1.config(text = self.counter1_txt)
        chk2.config(text = self.chk2_txt)
        counter2.config(text = self.counter2_txt)
        com.config(text = self.com_txt)
        ent.config(text = self.ent_txt)
        chk3.config(text = self.chk3_txt)
        
    def load_data(self, filename, input_sheet, column_names, start_row, max_row):
        print('Now translating' + str(filename))
        modulepath = os.path.dirname(xw.__file__)
        print(modulepath)
        
        col = ','.join(column_names) #joins column names as a string with , - needed to set the columns for the dataframe
        rows = int(int(max_row) - int(start_row)) #number of rows is max_row minus start row - b/c max row refers to all of the filled 
        self.rows = rows
        
        self.ldat_txt = 'The length of data is ' + str(rows) + ' rows. The starting row is ' + str(start_row) +'\n'
        print(self.ldat_txt)
        
        
        
        skipline = int(start_row) - 1 #this is done b/c pd starts with 0-index so 0 = 1st line, etc. and we want to skip the lines before the start row (so if starting row is 5 we want to skip 4 lines (0,1,2,3)
        data = pd.read_excel(filename, input_sheet, header=None, names=column_names, usecols=col, skiprows = skipline, nrows = rows) #looked at pandas documentation for how to work with excel
        
        if len(column_names) > 1:
            data[column_names] = data[column_names].astype(str) + '. '
            data["combined"] = data.sum(axis=1)
            untranslated = data["combined"].tolist()
        else:
            untranslated = data[column_names[0]].tolist()
        self.untranslated = untranslated
        
        names = pd.read_excel(filename, input_sheet, header=1, usecols = col) #creates a separate dataframe for the names of columns (which are not grabbed in the original dataframe due to the starting row being at the data, rather than the columns
        self.names = names #sets it so can be referenced outside of the function
        #print(names.columns) #as a check - it works
        self.extr_txt = 'Data extracted'
        print(self.extr_txt)
        print(self.untranslated)
        
    def line_by_line(self):
        '''A function to make the program quicker by calling the data check translation and input functions
        for each line rather than doing it all at once'''
    
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
        print(str(ilang) + ' ' + str(olang))
        
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
        for un in self.untranslated: #for each string (un) in untranslated
            if str(un) == '' or str(un) == None:
                self.checked_untrans.append(un)
                n=n+1
                self.counter1_txt = str(n) + ' entries checked'
                print(self.counter1_txt)
            elif str(un) == 'nan':
                self.checked_untrans.append('N/A')
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
                        un = check_un.lower().replace(pats[i], reps[i]) #un becomes the lowercased string, with the pattern replaced
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
                    self.checked_untrans.append(caps) #and appended to checked_untrans list for translation
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
                    self.checked_untrans.append(fin) #appends to checked untranslated
                    n=n+1
                    self.counter1_txt = str(n) + ' entries checked'
                    print(self.counter1_txt)
                
    def translator (self, ilanguage, olanguage):
        
        self.chk2_txt = 'Now translating - this can take a while, make a cup of tea :)'
        print(self.chk2_txt)
        try:
            n=0
            for i in self.checked_untrans: #for strings in untranslated
                
                transl_text = str(i) #converts it into a string (to be sure that it is actually coming out as string value)
                transltxt = transl(source=ilanguage, target=olanguage).translate(transl_text) #translates the string - taken from deep_translator documentation#
                self.translated.append(transltxt)
                n= n+1
                self.counter2_txt = str(n) + ' entries translated'
                print(self.counter2_txt)
                 #so progress can be checked - and can see if it gets stuck on any particular entries
                #print(transltxt) #optional - can uncheck to see the translated results
                #print('translating2')
                #print(i + '\n \n' + transltxt +'\n \n')
        except Exception as e: #except if there is an error 
            self.err_txt = str(traceback.print_exc())
            print(str(traceback.print_exc()))
            
            #print('''IF THE CONNECTION HAS TIMED OUT, WAIT SOME TIME/TRY AGAIN OR TRY PUTTING IN headers={
                                    #'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_9_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/35.0.1916.47 Safari/537.36'}
                                   # INTO deep translator's google_trans translate function in response = requests.get on line 117 ''')
 
        #print("If it's not proceeding, check if Excel has opened the file because it contains macros, click 'x' on the macros to access the spreadsheet and program will resume")    
            
    def combinedata(self, column_names, input_column):
        '''
        This function is for columns like Location Notes, where the English translation is input in the same entry underneath the Russian data rather than in its own column. So we need to combine
        the untranslated and translated data into one entry, which can then be input into the column
        '''
        combine = True
        self.combine = combine
        if input_column in column_names: #if the input_column is also in column_names aka if data is being extracted from the same column that it should be put back into - so there is both Russian and English in the entry
            self.com_txt = 'Combining data from ' + str(column_names)
            print(self.com_txt)
            
            for i in self.untranslated:
                if str(i).lower() == 'nan':
                    self.combineddata.append('N/A')
                else:
                    num = int(self.untranslated.index(i))
                    combined = str(i) + ' \n ' + str(self.translated[num])
                    self.combineddata.append(combined) #appends combined string to combinedddata
                #print(combined)
            
            combine = True #so if we want to be inputting the result to the same column that the original data was in without overriding the original data
            self.combine = combine #sets it so it can be referenced outside of the method
        else:
            combine = False #combine is set to false
            self.combine = combine #sets it so we can reference it outside of this method
            pass

        
    def input_data(self, filename, input_sheet, input_column, start_row, max_row):

        '''
        This function inputs the translated data from the list into the cells using xlwings as xw - it opens the Sheet, inputs the data using a combination of the input_column and given row, knowing that the data is in
        the same order because that's the order it was extracted in and worked on.
        
        The below code emerged from combined xlwings documentation and various tidbits of code obtained from StackOverflow as I was running into errors.
        '''
        
        wb = xw.Book(filename) #defines wb as calling Book method for filename in xw
        sheet=wb.sheets[input_sheet] #sheet accesses the sheets from the workbook (using the input_sheet) #allows user to input column (varies based on
        #query and spreadsheet)
        self.ent_txt = 'Entering data into spreadsheet'
        print(self.ent_txt)
        
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
        
    def runn(self, filename, input_sheet, column_names, input_column, start_row, max_row, ilanguage, olanguage, filetype, glossfile): #the whole translation and SQL input process with the associated variables
        self.transl.load_data(filename, input_sheet, column_names, start_row, max_row)
        self.transl.data_check(ilanguage, olanguage, filetype, glossfile)
        self.transl.translator(ilanguage, olanguage)
        self.transl.combinedata(column_names, input_column)
        self.transl.input_data(filename, input_sheet, input_column, start_row, max_row)
    
    def SQLinput(self, filename, folder, input_sheet, tables, data_cols, max_row):
        self.transl.sql_work(filename, folder, input_sheet, tables, data_cols, max_row)
        
    
    def runtransl(self, filename, input_sheet, column_names, input_column, start_row, max_row, ilanguage, olanguage, filetype, glossfile):

        print('Now translating: ' + str(filename) + '\n Columns: ' + str(column_names) +'-->' + str(input_column))
        print('Note: If file contains forms/macros you will need to manually close/log into these when they pop-up during entry')
        
        
        self.runn(filename, input_sheet, column_names, input_column, start_row, max_row, ilanguage, olanguage, filetype, glossfile)
        self.transl.reset()
    
        