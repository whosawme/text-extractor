# text-extractor
Extract text from multiple Word documents in a folder, and parse the identified sections into separate columns of an excel spreadsheet output.

Script will go through all .docx files and copy text starting from each start_word to its stop_word or conditional_word variables. 
The body of text between the start and stop words will comprise of each column in the excel output. 
Configure your start and stop words identifying the body of text for each column at the top of the script, as well as the columns of your output file. 


Required Modules:
os
pandasdocx
re
docx
xlsxwriter
