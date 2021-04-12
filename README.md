# sqlite
 
This is a simple function to populate an existing word template with values from a sqlite database.
Modify it for your own needs. 

An example is included in the repository. If you run "beispiel.py", then a connection is established to the database "beispiel.db". Then an sql-query is ran and the results of the query are fetched and then compared to the merge fields within the word document "beispiel.docx". Lastly, a different word document "beispiel_output.docx" is created within the same folder in which the merge fields have been populated with the relevant values from the database.
