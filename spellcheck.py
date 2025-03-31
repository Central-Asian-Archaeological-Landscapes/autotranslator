'''
Created on 19 Jan 2022

@author: turch
'''
import enchant
import enchant.checker
import os
import re
import difflib
import pandas as pd
import spellchecker as sc

class Spelling():
    newstring = None
    
    def ru_dict(self):
        '''
        A short function to move dictionary files from the data folder in src code to the pyenchant folder
        on the user's drive so that pyenchant can then use it as a Russian dictionary. This is done because although
        I could move the files manually to my own pyenchant folder, it needs to also be done for any user. 
        It is simply done by finding the paths to the dictionary files and to the module, and then renaming the file path 
        which moves the files to the required directory. The files come from LibreOffice and are open access, I was able 
        to get them by downloading the software and looking in its language folders using advice from a Russian article
        on how to write a spell-checker.
        '''
        pn = os.path.split(os.path.dirname(os.path.abspath(__file__)))[0] 
        data = os.path.join(pn, 'src', 'data', 'dict', 'dictionary files')
        files = os.listdir(data)
        
        modulepath = os.path.dirname(enchant.__file__)
        folder = os.path.join(modulepath, 'data', 'mingw64', 'share', 'enchant','hunspell')
        print(folder)
        for file in files:
            file_path = os.path.join(data, file)
            new_path = os.path.join(folder,file)
            if os.path.exists(new_path):
                pass
            else:
                os.rename(file_path, new_path)
    
    def create_glossary(self, ilang): 
        
        '''
        here we need to create a dictionary which the untranslated data can be compared against using the thesaurus spreadsheet developed by CAAL. The ditionary is created by using pandas to read the spreadsheet
        and make a dataframe of it, converting it to a CSV file as for some reason the Russian here did not want to encode otheriwse, and creating a dictionary with the Russian words set as keys and the English words
        as values; this means that strings can then be compared against the keys to see if they are present, and then substitute the keys for the values.
        
        '''
        
        project_folder = os.path.split(os.path.dirname(os.path.abspath(__file__)))[0] #splits path from the directory name to the main module being opened here
        data_folder = os.path.join(project_folder, "src") #joins the main module path with that of data folder so it can be accessed
        glossfile = os.path.join(data_folder, 'data', "glossary.xlsx") #joins the data folder path with image folder so it can be opened/accessed
        
        #various StackOverflow queries helped with understanding errors that were arising here, particularly in terms of encoding
        if ilang == 'ru':
            
            rus = pd.read_excel(glossfile, usecols = 'B:C')#uses pandas to read excel file extracting columns B to C (Russian and English)
            rus.to_csv("Rusgloss.csv", index = None, header=True) #converts excel to csv because otherwise there are encoding issues that cannot be addressed in pandas read_excel because it removed ability to specify encoding
            rusdf = pd.DataFrame(pd.read_csv("Rusgloss.csv", encoding = 'utf-8')) #reads created csvfile to dataframe specifying the encoding so the Russian characters are read properly
            russgloss = rusdf.set_index('Russian').T.to_dict('list')
            #print(russgloss)
            self.russgloss = russgloss #sets it so it can be referred
            return russgloss
            print('Glossary created')
        elif ilang == 'zh_CN':
            cn = pd.read_excel(glossfile, usecols = 'C:D')
            cn.to_csv('Chinese_gloss.csv', index = None, header = True)
            cndf = pd.dataframe(pd.read_csv('Russgloss.csv', encoding = 'utf-8'))
            cngloss = cndf.set_index('Chinese').T.to_dict('list')
            self.cngloss = cngloss
            return cngloss
        
    def xl_spellcheck(self, string, ilang):
        '''A function to check each word for misspellings by comparing it with the pyenchant spellchecker using the russian dictionary files as d. 
        This function is applied to each untranslated string within translmethods.
        To operate the spell-checker, we first need to remove proper nouns as they will inevitably come back as misspelled and be erroneously replaced. This is done by using re to strip the string
        of any words beginning with a capital letter - which I know will also omit nouns at the start of a sentence so a better method will be implemented in the future.
        Then words in Russian quotes are also removed as these relate to names or Central Asian terms which are better off not being corrected in the first place
        And then punctuation . and , are removed to make analysis easier
        And double spaces are set to single spaces and trailing spaces are removed again '''
        
        global newstring #makes newstring a global variable so it can be accessed from translate methods
        newstring = string
        
        latinalph = ['fr', 'uz', 'en']
        #stackOverflow queries by other users assisted with understanding regex
        if ilang == 'ru':
            
            nstring = re.sub(r"\s*[А-Я]\w*\s*", " ", string).strip() #removes all words beginning with capital letters and removes spaces at the start/end
            nstring = re.sub(r'(?s:(?=(?P<g0>.?«))(?P=g0)(?=(?P<g1>.*?»))(?P=g1).*)\Z', '', nstring) #uses regex to remove any words within russian parentheses - which includes Central Asian terms and names
            
        elif ilang in latinalph:
            nstring = re.sub(r"\s*[A-Z]\w*\s*", " ", string).strip()
        
        else:
            pass
           
       
        nstring = re.sub(r'(?s:[.,])\Z', ' ', nstring) #removes punctuation
        
        nstring = nstring.replace('  ', ' ').strip() #replaces double spaces with single ones, and if there are still spaces at start/end of string, removes those
        
        chkr = enchant.checker.SpellChecker('ru_RU', nstring)
       
        if ilang == 'uz':
            return newstring
        elif ilang == 'zh_CN':
            return newstring
        elif ilang == 'kz':
            return newstring
        
        d = enchant.Dict('en_GB') #sets the english dictionary so it can be accessed - so if english words pop up they're not marked as errors
    
        '''
        The below code was adapted from a StackOverflow question on how to automate spell-check/auto-corrector. I added the if d.check == True to catch out English words because I am working
        with Russian data, and added the if len > 0 because single characters were getting marked as error-words when they are actually initials - but the rest of the code stayed the same
        '''
        for err in chkr:
            if d.check(err.word) == True: #sometimes english letters/words might appear in the strings and are raised as errors because it's set to Russian
                continue
            elif len(err.word) > 1: #if the length of the erroneous word is more than one character (i.e. it is not an initial or wayward letter which if corrected here will be made more erroneous
                #print('error word:' + str(err.word)) to check which words are being marked - for me to improve the dictionary/thesaurus further
                if len(err.suggest()) > 0: #if there are suggestion words (i.e. the list is not empty)
                    sug = err.suggest()[0] #take the first word (WILL BE OPTIMISED IN FUTURE USING DIFFLIB SEQUENCE MATCHER RATIOS TO FIND BEST MATCH OF THE WORDS - but this may involve some machine learning)
                    newstring = newstring.replace(str(err.word), str(sug)) #newstring is newstring where the erroneous word is replaced by the suggested substitute
                else: #otherwise if there are no suggested words
                    continue #continue on (likely not a mistake word but one not included in dictionary)
            else: #otherwise continue on
                continue
        
        #print('checked ' + str(newstring) + '\n') #test
        return newstring #newstring with replacements of checked words is the return result 

    def doc_spellcheck(self, ilang, word):
        d = enchant.Dict('en_GB')
        sim = dict()
        weird =['{','-', '_', '+', '}', '=', '<', '>', '/', '$', '£', ':', ';', '&', '%', '"',',','.' ]
        pn = os.path.split(os.path.dirname(os.path.abspath(__file__)))[0] 
        data = os.path.join(pn, 'data', 'dict')

        
        if d.check(word) == True:
            return word
        elif word[0] in weird:
            return word
        elif word[-1] in weird:
            word = word[:-2]
        else:
            pass
        
    
            
        if ilang == 'zh_CN':
            big_cn_gloss = os.path.join(data, 'TSAIWORD (Mandarin).txt')
            cn_gloss = r'' + big_cn_gloss
            big_gloss = open(cn_gloss, 'r', encoding = 'utf-8')
            gloss = str(big_gloss.read().split('\n'))     
        elif ilang == 'uz':
            return word
        else:
            return word
            
        if word in gloss:
            print('found')
            return word
        else:
            return word
            '''
            if r.check(word) == False and word not in gloss:
                suggestions = set(chkr.suggest(word))
                if len(word) > 1:
                    if len(suggestions) == 1:
                        sug = list(suggestions)[0]
                        print('sug1 ' + sug)
                        return str(sug)
                    elif len(suggestions) == 0:
                        print('len0 + ' + word)
                        return word
                    else:
                        print(list(suggestions))
                        for w in list(suggestions):
                            measure = difflib.SequenceMatcher(None, word, w).ratio()
                            sim[measure] = w
                            sug = sim[max(sim.keys())]
                            if sug == None:
                                sug = list(suggestions)[0]
                                print('sug2 ' + sug)
                                return str(sug)
                            else:
                                print('sug3 ' + sug)
                                return str(sug)
                else:
                    return word
            else:
                return word
            '''
    def spell_checker(self, filetype, ilang, word, russgloss):
        
        supported = ['en', 'es', 'fr', 'pt', 'de', 'ru']
        weird =['{','-', '_', '+', '}', '=', '<', '>', '/', '$', '£', ':', ';', '&', '%', '"',',','.' ]
        
        pn = os.path.split(os.path.dirname(os.path.abspath(__file__)))[0] 
        data = os.path.join(pn, 'src', 'data', 'dict')
        
        if str(word)[0] in weird and len(str(word)) == 1:
            return word
        elif str(word)[-1] in weird and len(str(word)) > 1:
            word = str(word)[:-2]
        else:
            pass
                
        if ilang in supported:
            spellcheck = sc.SpellChecker(ilang)
            
            if filetype == 'Word' or filetype == 'PDF':
                if ilang == 'ru':
                    russ_words = list(str(s) for s in russgloss.keys())
                    spellcheck.word_frequency.load_words(russ_words)
                    
                    d = enchant.Dict('ru_RU')
                    
                    big_ru_gloss = os.path.join(data, 'ru_dict.txt')
                    ru_gloss = r'' + big_ru_gloss
                    big_gloss = open(ru_gloss, 'r', encoding='utf-8')
                    gloss = str(big_gloss.read()).split('\n')
                    
                elif ilang == 'en':
                    d = enchant.Dict('en_GB')
                    gloss = []
                elif ilang == 'es':
                    d = enchant.Dict('es')
                    gloss = []
                elif ilang == 'fr':
                    d = enchant.Dict('fr')
                    gloss = []
                elif ilang == 'de':
                    d = enchant.Dict('de_DE_frami')
                    gloss = []
                
            
                if word in gloss or d.check(word) == True:
                    #print('found')
                    return word
                elif len(word) > 1:
                    misspelled = spellcheck.unknown(word)
                    if len(misspelled) > 0:
                            #print(spellcheck.candidates(word))
                            return spellcheck.correction(word)
                    else:
                            return word
                else:
                    return word   
                
            elif filetype == 'Excel':
                if ilang == 'ru':
                    russ_words = list(str(s) for s in self.russgloss.keys())
                    spellcheck.word_frequency.load_words(russ_words)
                    
                    #nstring = re.sub(r"\s*[А-Я]\w*\s*", " ", str(word)).strip() #removes all words beginning with capital letters and removes spaces at the start/end
                    nstring = re.sub(r'(?s:(?=(?P<g0>.?«))(?P=g0)(?=(?P<g1>.*?»))(?P=g1).*)\Z', '', str(word)).strip() #uses regex to remove any words within russian parentheses - which includes Central Asian terms and names
                    d = enchant.Dict('ru_RU')
                    
                    big_ru_gloss = os.path.join(data, 'ru_dict.txt')
                    ru_gloss = r'' + big_ru_gloss
                    big_gloss = open(ru_gloss, 'r', encoding='utf-8')
                    gloss = str(big_gloss.read()).split('\n')
                
                else:  
                    nstring = re.sub(r"\s*[A-Z]\w*\s*", " ", word).strip()
                    if ilang == 'en':
                        d = enchant.Dict('en_GB')
                        gloss = []
                    elif ilang == 'es':
                        d = enchant.Dict('es')
                        gloss = []
                    elif ilang == 'fr':
                        d = enchant.Dict('fr')
                        gloss = []
                    elif ilang == 'de':
                        d = enchant.Dict('de_DE_frami')
                        gloss = []
                    
                nstring = re.sub(r'(?s:[.,;!£$&";:?<>%])\Z', '', nstring) #removes punctuation
                nstring = nstring.replace('  ', ' ').strip() #replaces double spaces with single ones, and if there are still spaces at start/end of string, removes those
                split_word = nstring.split(' ')
                for wordd in split_word:
                    if wordd[0].isupper(): #if the first word is a capital ignore b/c proper noun
                        break
                    elif wordd in gloss or d.check(wordd) == True:
                        #print('found')
                        break
                    elif len(wordd) > 2:
                        if wordd not in gloss or d.check(wordd) == False:
                            checked_word = spellcheck.correction(wordd)
                            word = str(word).replace(wordd, checked_word)
                            break
                        else:
                            break
                    else:
                        break
                return str(word)
        else:
            if filetype == 'Word' or filetype == 'PDF':
                self.doc_spellcheck(ilang, word)
            elif filetype == 'Excel':
                self.xl_spellcheck(word, ilang)
            else:
                return word
        
            