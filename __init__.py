'''
Created on 19 Jun 2022

@author: turch
'''
import fullAutoTranslator
import threading

def init(self):
    GUI = fullAutoTranslator.GUI()
    threading.Thread(target=GUI.start).start() 

init(init)