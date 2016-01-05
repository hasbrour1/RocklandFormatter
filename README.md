# RocklandFormatter
Formats EDDs To Rockland's Excel Template Written in C#

This program allows the user to select an EDD excel file to format it to  the specifications of Rockland County's EDD's.

The main code is in Form1.cs.  The program opens the selected excel file along with the Rockland template excel file 
that is located in the same directory the exe was ran.  It copies each line from the selected file and puts it into the same
column in the template.  There are a few analytes that are skipped so each line is checked.  After each line is copied onto the 
template, the file is saved with the same name with _formatted at the end of the file in order to differentiate the final from the original.  
