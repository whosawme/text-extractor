# text-extractor
Extract text from multiple Word documents in a folder, parse the identified sections into separate columns of an excel spreadsheet output.

Define where each section starts and ends with start and stop words being the first word of a paragraph or newline. Script will go through all .docx files and copy text starting from each *start_word_x* variable to its *stop_word_x* or *conditional_word* variables. 
The body of text between the start and stop words will comprise of each column in the excel output. 
Configure your start and stop words identifying the body of text for each column at the top of the script, as well as the columns of your output file. 

Example:
start_word = 'Headline'
stop_word = 'Lessons'

Word doc sample:

bla bla bla
Headline  > bla bla bla bla

This text will be copied including the bla'blas above^
and will end here>.

Lessons:
this text wont be copied



to add more words tags add the coinciding if and while loop, as well as the column headers of the capturing DataFrame. 


Required Modules:
os
pandasdocx
re
docx
xlsxwriter




#whosawme
