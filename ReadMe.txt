Project Spell Checker 
Version 1.0
11/1/2000

Written by Greg DeBacker
gdebacker@windsweptsoftware.com

1) What It Does
2) Program Layout
3) Viewing The Log
4) Future Possibilities
5) Get The Latest Version
6) Version History

What It Does
Have you ever suffered the pain and embarrassment of a misspelled menu 
caption? If so, then this is the program for you. I don't consider myself 
to be a really bad speller but I make stupid mistakes. Transposing "i"s 
and "e"s, or maybe adding an extra "l" where there shouldn't be. I wrote 
this program to go through the entire project and extract every quoted 
string and every caption, text, and list property. It also will extract 
comments if you want. It then stores every thing in a ListView control and 
runs a spell checker on it. The spell checker comes courtesy of the Word 8 
Object, which means you will need MS Word 97 or greater to use it. Even if 
you don't have access to MS Word you can still review all of the words in 
the ListView. After the spell checker has been run the results of only 
those items that have been corrected are output to a log file that can 
be viewed in a text editor.

The program also gathers a lot of information about the project along the 
way, like how many lines of code, what objects, references, and DLLs are 
used, along with some other stuff.


Program Layout
On the left is a TreeView with a list of all of the project files, and 
information about each file. The Key property of each node in the 
TreeView holds the path to each file. The program uses this path to open 
the files. Click on the node to view the path information for each file.

On the right is a ListView that will display all of the string properties 
and variable assignments. The ListView columns are laid out like this:

File| Control/Routine| Type| Variable/Property| Value/Caption| Corrected Text

File - The name of the file the text was found in.
Control/Routine - The name of the control or the name of the sub or function.
Type - Is it a Label, TextBox, Sub, Function, etc.
Variable/Property - Is it a Caption, Text, List, or the name of the variable.
Value/Caption - The current value of the variable or property.
Corrected Text - Any corrections will show up in this column.

You will undoubtedly get more than you want because not all quoted strings 
in a program are actually words that the end user will see. Rather than 
miss something I decided to err on the side of caution and just grab 
everything that was a quoted string.

The program is smart enough to go inside FRX files and get the ListBox and 
ComboBox List property assignments. Sometimes this works perfectly and other 
times you get more than you need. Keep and eye out for this. It is also 
smart enough to remove the ampersand from captions with a hot key 
assignments before it sends them to the spell checker. If the caption is 
not spelled correctly then the program will try to put the ampersand back 
in the correct position on all suggested words. You should watch for this, 
though, to make sure it gets back the way it was.


Viewing The Log
After the spell checker is run, items in the ListView that have had 
corrections added to the last column will have a check mark next to them. 
Only those items with check marks will be output to the log file. Before 
you view the log file you can uncheck any items you don't want to view in 
the report.


Future Possibilities
This program could be expanded on to have it check for dead or redundant 
code. It could also check for functions that return variants because you 
for got to declare a type. Any number of things could be added. I thought 
about adding an option to have it update the project with the corrected 
spelling. Maybe when I get more time.

If you feel up to the task of adding a new module to this program or 
refining the current spell checker I encourage you to do so. If you 
downloaded this from somewhere other than my web site you should first go 
to the URL below and get the current version to see if anyone has improved 
upon it yet. Then, if you want to add to it write me an email and let 
me know what you're working on. When you've created a new module I'll add 
it to the program and post a new version on my web site and around other 
sites on the web.

Get The Latest Version
Download: http://windsweptsoftware.com/dl/vbspell.zip
Email Me: gdebacker@windsweptsoftware.com

Version History
Version 1.0
First released on 11/1/2000


Thank you,

Greg DeBacker
http://windsweptsoftware.com
