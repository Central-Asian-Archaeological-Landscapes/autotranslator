'''
Created on 8 Dec 2021

@author: SZXZ8
'''
import re
from deep_translator import GoogleTranslator as transl
from dataclasses import replace
import sqlite3 as lite
import sys
import pandas as pd
            
class SQLite():   
    '''
    This code was more or less directly taken from my finished Assignment 2 with some modifications; combining the create table and columns functions into one, and combining some aspects of 
    create_database, and removing some functiond which were not needed, and changing the variables somewhat to fit with these functions being called in a different module than before.
    '''
    def __init__(self, folder):
        self.folder = folder
        
        con = None
        
        try:
            self.con = lite.connect(self.folder)
            cursor = self.con.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
            print("Opened database " + self.folder) 
            #prints results from having the database (self is used b/c this method will be called from outside this class)
            print("\nContains tables: " + str(cursor.fetchall())) 
        except lite.Error: #except if there is an error - error is reported as %s (string) and quits
            ''' If it was not possible to open the database, report the error and quit. ''' 
            print ("Error %s:") #as above - prints error as %s (string)
            sys.exit(1)
            
    def close(self):
        if self.con: # if/else statement
            self.con.close()#if self.con is able to be closed
            print ("\nClosed database " + self.folder) #print closed database with filename
        else: #otherwise print that the database was not open
            print ("\nDatabase " + self.folder + " was not open!")

    def makesql(self, table, columns):
        sql_string1 = "DROP TABLE IF EXISTS %s" % table
        sql_string2 = "CREATE TABLE %s (%s)" % (table, columns)
        
        try: #try/except clause 
            self.con.text_factory = str #not sure what this is
            cur = self.con.cursor() #defines cur as the cursor within database connection - allows to alter database
            cur.execute(sql_string1) #executes sql_string1 where table DROPPED IF EXISTS
            cur.execute(sql_string2) #executes sql_string2 where NEW TABLE CREATED
        
                        
            self.con.commit() #commits changes
            '''tries to execute the strings defined above - aka this is the function that drops existing table, 
            then creates a new one with the required columns '''
        except lite.Error as e: 
        #If the table is not being created "roll back to previous state" is printed as error message
            ''' Or, if there was an error, roll back to the previous state '''
            if self.con:
                self.con.rollback() #goes back to where it was before it encountered the error
                print ("Error %s\nRolled back to previous state") % e.args[0] #
                
    def insert_row(self, table, data): 
        sql_string = "INSERT INTO %s VALUES(" % table #insert into table values
        for i in range(1, len(data)):
            #for i in the range of 1 to the length of the data (i.e. the number of columns in the file)
            sql_string += "?," #adds another value to variable's value and assigns that new value to variable
            # in this case adds an ? for each column to the end of the string
        sql_string += "?)" #then this adds the last '?' 
        #now the sql_string just needs to be combined with the actual values from the csvfile which will be done in create_database
        print(sql_string) #prints string so can see what is passed on
     
        ''' Now execute the SQL statement.  Note that you need to pass the actual data as the
            second argument to execute '''
        try: #try loop to exit database safely if there is error
            self.con.text_factory = str 
            '''text_factory controls which objects will be returned for TEXT parameter - in this case it is strings
            self.con is the connection to database, text_factory is aspect of con aka the SQL directory'''
            
            cur = self.con.cursor() #defines cur as cursor in SQL directory to alter database
            
            cur.execute(sql_string, data)
            '''executes sql_string + data so that it mimics the format
            "INSERT INTO table VALUES ("?", "?", etc.), ("3000 BC", "stone", etc.) - sql_string is the 
            question mark place-holders and data represents the values '''
            
            self.con.commit()   #commits the changes to the database
            ''' This will tell us if any rows have been updated '''
            print("Number of rows updated: %d"%cur.rowcount) 
            #prints number of rows updated as the method is applied to each row in the for loop in create_database
           
        except lite.Error as e: #except clause for the try statement - if there is error it goes back
            ''' Or, if there was an error, roll back to the previous state '''
            if self.con: #if error is within the database (con = connects to SQLite)
                self.con.rollback() #rollback and print error
                print ("Error %s\nRolled back to previous state"%e)
    
        
    def update_value(self, table, column, column2, value, key): #self is used because this is a class method
        #the other values are needed to
        '''
            This module simply updates a table in the database by selecting a
            single entry and updating the associated columns with new data 
        '''    
        
        ''' Assemble the SQL string '''
        sql_string = "UPDATE %s SET %s=? WHERE %s=?" % (table, column, column2) 
        #sql_string Updates the specified table to set column = ? where ID =?
        
        ''' Execute the SQL statement, passing the relevant data '''
        try: #try loop for editing SQL database
            cur = self.con.cursor() #cur defined as cursor allowing to edit database
            cur.execute(sql_string, (value, key))
             #the value and key are now called for the ? placeholders
             #the cursor is sued to execute the string
           
        except lite.Error:
            ''' Or, if there was an error, roll back to the previous state '''
            if self.con:
                self.con.rollback()
                print ("Error %s\nRolled back to previous state")    
                
            
    def analysis(self, table, SQLcolumns):
        att1 = 'Name'
        terms = ['"%kurgan%"', '"%burial ground%"', '"%mausoleum%"', '"%mosque%"', '"%hillfort%"']
        con = lite.connect(folder)
        cur = self.con.cursor() #cur is the cursor within con (the sqlite connection to database)
        
        for term in terms:
            stringg = "SELECT %s FROM %s WHERE %s LIKE %s", (SQLcolumns, table, att1, term,)

            cur.execute(stringg) #executes the sql_string given
            self.con.commit()   #commits the changes to the database
            ans = cur.fetchall() #cursor action of fetching all results from the query in the sql_string
            print('There are ' + str(len(ans)) + ' ' + str(term[2:-3]) + 's. \n')
            
            
                
            

