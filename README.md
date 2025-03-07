# NOTE:
1. The merge_row_cells() function may not function completely function for the
    'Instructors' Sheet, specifically the 'Staff' column. This is due to the fact
    that 'Staff' might overlap times as no proffesor is specified.
2. The main difference between compiled.py and remake.py is the
    conflict identification. This can be seen at the bottom of the
    'List View' sheet in finalOutput.xlsx
    - Another difference is that compiled.py converts (CSV) --> xlsx sheet
        and remake.py converts a (list of dictionaries) --> xlsx sheet

# USE:
1. Take the desired (.csv) file, in our case its called (finalSheet.csv), and put
    put it in the same folder as the (compiled.py) program file
2. Run the command to generate the current (MFW Labs), (TR Labs),
    (Instructor), and (List View) sheets in the file (output.xlsx):
        > python3 compiled.py
3. Because its a (.xlsx) file, you must use Microsoft Excel to open and modify.
4. After any modifications have been made on (output.xlsx) in the (List View)
    sheet, make sure the file is located in the same folder as the (remake.py)
    program file.
5. Run the following command to remake the (output.xlsx) file with the 
    desired modifications into the file (finalOutput.xlsx):
        > python3 remake.py
6. Open (finalOutput.xlsx) and go to the (List View) sheet. At the bottom of
    the sheet, you can find 2 lists:
    1. The first list, (Place Conflicts), contains any class where the
        professor is scheduled to be in 2 places at once
    2. The second list, (Person Conflicts), contains any class where 2
        professors are scheduled to be in the same place.


# IMPROVEMENTS:
1. Parameterize Code: There is a decent amount of repeated code that could 
    be parameterized better to reduce redundant lines of code
2. Input / Output Specifying: Currently the names of the input and output 
    documents are constant. Make a method to specify document names
    during the function call
3. Program Combination: There are currently 2 programs, (compiled.py) and 
    (remake.py). There should be a method to combine these programs into 
    one functioning program so you don't need to call 2 seperate program.
4. Function Organization: Currently, there are many functions in different
    locations. There should be fewer, or more clearly defined, imported files.