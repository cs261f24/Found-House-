# Developer Documentation


Environment Setup:
(Note: this assumes you already have Python up and running on your machine.)
1. To ensure smooth setup and avoid conflict with other libraries, we recommend setting up a virtual environment. 
    a. A word of warning: virtual environments created in an Ubuntu/WSL terminal can't be accessed/run from the terminals
       built into VSCode and vice versa. If you are doing your coding in VSCode, we highly recommend that you create your
       virtual environment in VSCode. 
2. Once you have your virtual environment up and currently running, use the command 
```
pip install -r requirements.txt
```
in your terminal to install all the supplementary libraries.


The Important Files Where Most of the Program is Contained:
A. [Main.py](Main.py) - This is the main script. It calls other functions and integrates all the functionalities into one place where
   everything can be run. 
B. [GUIexpirmentation.py](GUIexperimentation.py) - This contains all the code that creates the GUI and handles its events. 
C. [rows.py](rows.py) - This is where the functions to add and remove columns and rows reside. We were still in the process of attempting
   to integrate it into the GUI when we ran out of time, so... 
    a. In the main branch on GitHub, it is designed to work in the terminal and is not connected to the GUI at all yet. 
    b. The frazfinal branch is where the incomplete attempt to integrate rows.py into the GUI currently resides. It is not
       complete, and thus crashes fatally when run. 
D. [read.py](read.py) - This contains all the functions for searching for data in the Excel file and returning the cell that you are
   looking for. 
E. [openpyxl](openpyxl/) - This is the folder containing all the code in the openpyxl supplementary library.  
F. [FoundHouse.xlsx](FoundHouse.xlsx) - This is the training Excel workbook, containing the fake data we were given so we could test our
   program. Its file path is hardcoded into our code, so that might need to be changed down the line. 
G. [Attribute.py](Attribute.py), [Owner.py](Owner.py), and [Pet.py](Pet.py) - These are leftover files from when we were considering an object-oriented approach.
   We're keeping them here just in case they're useful down the line. 
    a. Attribute objects were going to be a specific property of each pet - analogous to the columns in the Excel sheet. 
    b. Pet objects were going to be collections of information about each pet - analogous to the rows in the Excel sheet. 
    c. Owner objects were going to be collections of information about each owner - similar to Pet objects, but they were
       going to be distinct, as Pets and Owners would have different Attributes. 