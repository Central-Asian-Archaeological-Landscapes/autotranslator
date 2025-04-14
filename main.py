'''
Created on 19 Jun 2022

@author: turch
'''
from set_parameters import Begin
import threading

def init(self):
    global data_path
    data_path = input("Enter data path:")

    ilanguage = input ("Enter language of data - Options: 'ru', 'en':")

    olanguage = input("Enter language to be translated into:")

    filetype = 'Excel'
    run.run(self, data_path, ilanguage, olanguage, filetype)
    
init(init)