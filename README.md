# text-extractor
Extract text from multiple Word documents in a folder, parse identified sections into separate columns of an excel spreadsheet output.

Define where each section starts and ends with start and stop words being the first word of a paragraph or newline. Script will go through all .docx files and copy text starting from each *start_word_x* variable to its *stop_word_x* or *conditional_word* variables. 
The body of text between the start and stop words will comprise of each column in the excel output. 
Configure your start and stop words identifying the body of text for each column at the top of the script, as well as the columns of your output file. 

<h3>Example:</h3>
<h5>if:</h5>
  
start_word = 'Startword' 

stop_word = 'Stopword'


<b> then in Word doc sample: </b>

this text wont be copied

°Startword  bla bla bla 

This text will be copied including the bla'blas above^.

°Stopword:
this text wont be copied

<br>
<br>
<b>note:</b><br>
to add more words tags add the coinciding if and while loop, as well as the column headers of the capturing DataFrame. 
<br>
<br>
<br>
<b>Required Modules:</b><br>

os<br>
pandasdocx<br>
re <br> 
docx <br> 
xlsxwriter <br> 
<br>
<br>

#whosawme
