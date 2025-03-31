'''
Created on 3 Feb 2022

@author: turch
'''
import tkinter as tk
from tkinter import *
from src.run import run
from deep_translator import GoogleTranslator as transl
'''
The goal here is to create a GUI for the Translator program with user input. First the start window is created welcoming the user and prompting them to click Start. The data_path window then appears and it is input. Once 
the data path is put in, the openfolder function can be called from run
'''

class GUI:

    paths = 0 #creates path variable
    input_language = 'ru'
    output_language = 'en'
    filetype = 'Excel'
    
    def start(self):
        global window
        window = tk.Tk()
        frame = tk.Frame(master=window, width = 150, height=150)
        frame.pack() #makes a frame and packs it to put widgets within
        
        def enter_path(): #closes existing window and opens the set path
            frame.destroy()
            GUI.path_window(self)
            
        
        
        greeting = tk.Label(master=frame, text="Welcome to the CAAL Auto-translator for all your auto-translation file needs. \n \n Make sure needed files are closed before starting, the program will not work if they are open")
        start_btn = tk.Button(master=frame, text = 'START', command=enter_path) #
        greeting.pack(fill = tk.X)
        start_btn.pack(fill = tk.X)
        window.mainloop()
    
    def runrun(self):
        #lf = loadFolders()
        pinput = GUI.paths #takes the first entry of paths
        window.destroy()
        #return str(pinput), str(GUI.input_language), str(GUI.output_language), str(GUI.filetype)
        #lf.run(self, pinput, str(GUI.input_language), str(GUI.output_language), str(GUI.filetype))
        
        
    def path_window(self): #codes a window for entering data path, setting language of data and to translate into
        frame = tk.Frame(master=window, width = 300, height=300)
        frame.pack(fill = tk.X, side = tk.TOP)
        
        path = tk.Label(master=frame, text = 'Enter data path:')
        path_input = tk.Entry(master=frame)
        path.pack(fill = tk.X, side=tk.LEFT)
        path_input.pack(fill = tk.X, side = tk.RIGHT)
        
        framea = tk.Frame(master=window, width=150, height=150)
        framea.pack(fill = tk.X, side =tk.TOP)
        
        langs_dict = transl.get_supported_languages(as_dict = True)
        lang_list = list(langs_dict.keys())
        abb_list = list(langs_dict.values())
        
        print(langs_dict)
        
        def input_lang(selection):
            position = lang_list.index(str(ilang.get()).lower())
            GUI.input_language = abb_list[position]
            print(GUI.input_language)
            
            
        def output_lang(selection):
            position = lang_list.index(str(olang.get()).lower())
            GUI.output_language = abb_list[position]
            print(GUI.output_language)
            
        def fltype(selection):
            GUI.filetype = ftype.get()
        
        ilang = StringVar(framea)
        ilang.set("Russian")
        ilanguage = OptionMenu(framea, ilang, 'Russian', 'English', 'Chinese (Simplified)', 'Kazakh','Kyrgyz', 'Tajik', 'Turkmen', 'Uzbek', 'Italian', 'French', 'Spanish','German', 'Arabic', command = input_lang)
        lang = tk.Label(master=framea, text = 'Enter language of data, set to Russian by default:')
        lang.pack(fill = tk.X, side = tk.LEFT)
        ilanguage.pack(fill = tk.X, side = tk.RIGHT)
        
        frameb = tk.Frame(master=window, width=150, height = 150)
        frameb.pack(fill = tk.X, side = tk.TOP)
        
        lang2 = tk.Label(master = frameb, text = 'Enter language to be translated into, English as default:')
        lang2.pack(fill = tk.X, side = tk.LEFT)
        olang = StringVar(frameb)
        olang.set("English")
        olanguage = OptionMenu(frameb, olang, 'Russian', 'English', 'Chinese (Simplified)', 'Kazakh','Kyrgyz', 'Tajik', 'Turkmen', 'Uzbek', 'Italian', 'French', 'Spanish','German', 'Arabic', command = output_lang)
        olanguage.pack(fill = tk.X, side = tk.RIGHT)
        
        framec = tk.Frame(master=window, width =150, height = 150)
        framec.pack(fill = tk.X, side = tk.TOP)
        
        typefile = tk.Label(master=framec, text = 'Select type of file to be translated, Excel as default')
        typefile.pack(fill = tk.X, side = tk.LEFT)
        ftype = StringVar(framec)
        ftype.set('Excel')
        filetypes = OptionMenu(framec, ftype, 'Excel', 'PDF', 'Word', command = fltype)
        filetypes.pack(fill = tk.X, side = tk.RIGHT)
        
       
        def next_window(): #defines next_window function
            GUI.paths = str(path_input.get())
            GUI.runrun(self) #runs
            
        framec = tk.Frame(master=window, width = 150, height = 150)
        framec.pack(fill = tk.X, side = tk.TOP)    
        ok_btn = tk.Button(master=framec, text = 'OK', command=next_window)
        ok_btn.pack(fill = tk.X, side = tk.BOTTOM)
        
        
        
        window.mainloop()



