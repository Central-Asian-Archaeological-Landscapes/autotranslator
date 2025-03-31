'''
Created on 30 May 2022

@author: turch
'''

import tkinter as tk
import sys
from tkinter import *

class ResetGUI:
    def restart(self):
        window = tk.Tk()
        framek = tk.Frame(master=window, width = 150, height = 150, padx=5, pady=5)
        framek.pack(fill = tk.X, side = tk.TOP)
        framei = tk.Frame(master=window, width=150, height = 150, padx = 5, pady=5)
        framei.pack(fill=tk.X, side = tk.TOP)
        rs = tk.Label(master=framek, text = 'Would you like to restart with another data path?')
        rs.pack(fill = tk.X, side = tk.TOP)
        
        def callback(selection):
            if str(selection) == 'Yes':
                window.destroy()
                self.reset()
                
            else:
                framek.forget()
                goodbye = tk.Label(master=framei, text = 'Exiting....Goodbye!')
                goodbye.pack(fill=tk.X, side = tk.TOP)
                window.after(3000, window.destroy)
                sys.exit()
            
        
                
        
        cont = StringVar(framek)
        cont.set('-')
        filetypes = OptionMenu(framek, cont, 'Yes', 'No', command = callback)
        filetypes.pack(fill = tk.X, side = tk.RIGHT)
        window.mainloop()
    def reset(self):
        from src.AutoTranslator import GUI
        GUI = GUI()
        GUI.start()