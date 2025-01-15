# CustomExcelFilter
Utilizing VBA code in excel, I designed a program that runs in the background of an excel file that allows for the use of interactive
filtering buttons. The set specific filtering buttons are tied to a table that scans through for specific parameters I set and filters 
to show the specifically selected condition. 

The buttons can be implemented in the excel file using the 'Developers' tab and inserting form controls in which the program will be linked to the form controls
via running macros. 

**IMPORTANT NOTE ON MACROS**
As I mention macros, many companies specific to mine have strict policies when running macros, and other programs on workplace machines. If you attempt to 
create this program via the VBA editor in excel (alt+F11 <--'to open VBA editor in excel workbook') change your privacy settings in excel and allow macros to run.
If further authentication is required reach out to your companies IT department and create a developers ticket. 

In creating this program, I also had to create a custom function within excel, called '=CountByColor' that was also created by utilizing VBA code
because '=CountByColor' is not an already specific function within excel. 

Some important snippets about this code is I can only share the specific program to protect sensitive company information stored in the excel file 
in which this program was originally written in. However, the code includes specific comments throughout the program explaining the reason for each section of code. 

