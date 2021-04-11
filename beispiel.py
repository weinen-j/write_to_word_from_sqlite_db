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

data = dict(zip(r.keys(), r))
print(data)

document = MailMerge(template)
#document.get_merge_fields()
document.merge(**data)
document.write("Beispiel-output.docx")

conn.close()