# -*- coding: utf-8 -*-
"""
Created on Sat Oct 24 18:45:45 2020

@author: Jonas
"""

#import dependencies
from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
import sqlite3


#establish connection to "beispiel.db"
conn = sqlite3.connect("beispiel.db")
conn.row_factory = sqlite3.Row
c = conn.cursor()

#define template
template = "Beispiel.docx"

#execute query
c.execute("SELECT * FROM applicants")
r = c.fetchone()
r.keys()

data = dict(zip(r.keys(), r)) #here you parse the data from DB
print(data)

#Second Version
#print(data.keys() = data.values())
data_dat = [] #create a list so we can populate the fields u set in word document ok!

document = MailMerge(template)
for i in document.get_merge_fields():      #here you parse through the fields u set in word and add to the list
    data_dat.append(i)                  
      
print(data_dat)

if all(val in data for val in data_dat): #this code is only to show how u can parse through a list, 
    print(val)

for j in data:          # in this we parse through the list and for each value we parse in the db 
    print(j, data[j])
str_str = []  
for i in range(len(data_dat)):
    for j in data:
        if data_dat[i] in j:
            #print(data_dat[i])
            str_merge = data_dat[i] + "=" + "'"+data[j]+"'"         # we create the format for merge()
            #print(data_dat[i], "=", "'",data[j],"'")
            str_str.append(str_merge)                           # we push each value 
print(str_str)
#Second Verion end


#passing only dictionary
document = MailMerge(template)
document.merge(**data)
document.write("Beispiel-output.docx")                      # with simple kwarg conversion, fastest

document.merge(str_str)
document.write("Beispiel-output-str.docx")                      # with str we created v check example

conn.close()