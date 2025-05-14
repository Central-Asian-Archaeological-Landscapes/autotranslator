'''
Created on 19 Jun 2022

@author: turch
'''
from set_parameters import Begin
import threading

def init(self):
    data_path = input("Enter data path:")

    ilanguage = input ("Enter language of data - Options: 'ru', 'en':")

    olanguage = input("Enter language to be translated into:")
    
    glossfile = input('Enter path for glossary (glossary.xlsx)').strip('"').strip()

    filetype = 'Excel'
    Begin.modrun(self, data_path, ilanguage, olanguage, filetype, glossfile)
    
init(init)