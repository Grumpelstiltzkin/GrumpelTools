"""
GrumpelTools
By Ian Holder
GrumpelStudios

Version 1.0 completed 18 Mar 2022
Version 1.1 Completed 2 Feb 2023
    Added GetExcelFile function
Version 1.2 Completed 6 Mar 2023
    Added Proper SciPy Documentation
Version 1.2.1 completed 7 Mar 2023
    Added type hinting
Version 1.2.2 completed 8 Mar 2023
    Fixed line lengths (<72 chars)
    Added OopsTryAgain()
    Added OopsPermission()
    Added BinarySearch()

INTENDED FUNCTION
A module that holds the methods I will want to use again in other
projects

KNOWN ISSUES
"""

VERSION = "1.2.2"

# Import modules
import tkinter
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import askquestion
import docx2txt
import datetime
import os
import openpyxl

#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-#

def banner(text: str, text1: str =""):
    """
    Print a banner with the input text inside the first line and text1
    inside the second line. Banner scales to fit the largest line of
    the two.

    Parameters
    ----------
    text : str
        the first line of text in the banner
    text1 : str, optional
        optional argument that will create a second line of text if
        provided

    Prints
    -------
    Each applicable line of text is padded with spaces to match lengths
    """

    if text1 != "":
        dif = len(text)-len(text1)
        while dif > 0:
            text1 += " "
            dif -= 1
        while dif < 0:
            text += " "
            dif += 1

    print("+--" + ( len(text) * "-" ) + "+")
    print("| " + text + " |")
    if text1 != "":
        print("| " + text1 + " |")
    print("+--" + ( len(text) * "-" ) + "+")
    print()

#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-#

# This prints a heading line with the argument string inside
def heading(string):
    print('~|', string, '|~')

#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-#

def OopsTryAgain():
    """
    Creates a question window with title: "Error"
    Displays the message:
        "Oops, that didn't work, try again?"
    Displays the buttons:
        yes
        no

    Returns
    -------
    yes
        if yes is clicked
    no
        if no is clicked
    """

    return askquestion(title = "Error", message = "Oops, " \
                "that didn't work, try again?")

#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-#

def OopsPermission():
    """
    Creates a question window with title: "Error"
    Displays the message:
        "Oops, you don't have permission to use that file.
        It may be open, close it and try again?"
    Displays the buttons:
        yes
        no

    Returns
    -------
    yes
        if yes is clicked
    no
        if no is clicked
    """

    return askquestion(title = "Error", message = "Oops, you don't " \
                       "have permission to use that file.\nIt may " \
                        "be open, close it and try again?")

#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-#

def GetTextFile(fileAskPhrase: str, initDir: str = ""):
    """
    Opens a Text file through an OS window input
    Validates that the file is the correct type

    Parameters
    ----------
    fileAskPhrase : str
        string used for the OS file system description
    initDir : str
        string of the directory to open in the OS window
        default is ""

    Returns
    -------
    file
        the opened file selected in the dialog
    """

    print("Getting text file...")
    while True:
        tkinter.Tk().withdraw()
        f = askopenfilename(title = fileAskPhrase,
                            initialdir = initDir)
        try:
            txtFile = open(f, 'r')
            return txtFile

        except PermissionError:
            answer = OopsPermission()
            if answer == 'no':
                break

        except:
            answer = OopsTryAgain()
            if answer == "no":
                break

#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-#

def GetWordFile(fileAskPhrase: str, initDir: str = ""):
    """
    Opens a Word file through an OS window input
    Validates that the file is the correct type

    Parameters
    ----------
    fileAskPhrase : str
        string used for the OS file system description
    initDir : str
        string of the directory to open in the OS window
        default is ""

    Returns
    -------
    file
        the opened file selected in the dialog
    """

    print("Getting text file...")
    while True:
        tkinter.Tk().withdraw()
        f = askopenfilename(title = fileAskPhrase,
                            initialdir = initDir)
        try:
            WordFile = docx2txt.process(f)
            return WordFile

        except PermissionError:
            answer = OopsPermission()
            if answer == 'no':
                break

        except:
            answer = OopsTryAgain()
            if answer == "no":
                break

#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-#

def GetExcelFile():
    """
    Opens an excel file through an OS window input
    Validates that the file is the correct type

    Parameters
    ----------
    fileAskPhrase : str
        string used for the OS file system description
    initDir : str
        string of the directory to open in the OS window
        default is ""

    Returns
    -------
    file
        the opened file selected in the dialog
    """

    print("Getting Excel file")
    while True:
        tkinter.Tk().withdraw()
        f = askopenfilename(title = "Select the excel file you want to open")
        try:
            wb = openpyxl.load_workbook(f)
            return wb

        except PermissionError:
            answer = OopsPermission()
            if answer == 'no':
                break

        except:
            answer = OopsTryAgain()
            if answer == "no":
                break

#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-#

def DateString():
    """
    Returns
    -------
    a string of the current date in format YYMMDD
    """

    return str( datetime.datetime.today().strftime("%y%m%d") )

#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-#

def DateTimeString():
    """
    Returns
    -------
    a string of the current date and time in format
        YYMMDD - HH:MM:SS
    """

    return str( datetime.datetime.now().strftime("%y%m%d_%H_%M_%S_") )

#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-#

def HamiltonDateString():
    """
    Returns
    -------
    a string for the current date in Hamilton Database file format
        YYYY-MM-DD HH-MM-SS-msZ.bak
    """

    return str( datetime.datetime.now().strftime(\
        "HamiltonXRP2_%Y-%m-%d %H-%M-%S-%fZ")[:-5]+"Z" )

#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-#

def DTCreateFile(strObj, strName, fileType=".txt", directory=""):
    """
    Creates a Save File and writes strObj to that file

    Parameters
    ----------
    strObj : obj
        a string object that needs to be saved as a file
    strName : str
        the string that names the file
    fileType : str, optional
        if provided, sets the .* file ending, defaults to .txt if not 
    directory : str, optional
        if provided, this is the direcory to create the file in
        defaults to current directory

    Returns
    -------
    str
        the generated str used as the filename
        this version includes DateTimeString as a prefix
    """

    print("Creating Date Time File...")
    fileName = DateTimeString() + strName + fileType
    if directory != "":
        fileName = os.path.join(directory, fileName)
    saveFile = open(fileName, "w")
    saveFile.write(strObj)
    saveFile.close()
    print('Created ' + fileName)
    return fileName

#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-#

def CreateFile(strObj, strName, fileType=".txt", directory=""):
    """
    Creates a Save File and writes strObj to that file

    Parameters
    ----------
    strObj : obj
        a string object that needs to be saved as a file
    strName : str
        the string that names the file
    fileType : str, optional
        if provided, sets the .* file ending, defaults to .txt if not 
    directory : str, optional
        if provided, this is the direcory to create the file in
        defaults to current directory

    Returns
    -------
    str
        the generated str to use as a filename
    """

    print("Creating File...")
    fileName = strName + fileType
    if directory != "":
        fileName = os.path.join(directory, fileName)
    saveFile = open(fileName, "w")
    saveFile.write(strObj)
    saveFile.close()
    print('Created ' + fileName)
    return fileName

#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-#

def BinarySearch(list, tgt):
    """
    Searches for the target in a list, assumes list is sorted in
    ascending order

    Parameters
    ----------
    list
        a list of potential targets to search for, can be any type
    tgt
        the thing you want to find, can be any type

    Returns
    -------
    found
        the index of tgt within the list
        OR
        -1 if not found
    """

    lo = 0
    hi = len(list) - 1
    found = -1

    while lo <= hi and found == -1:
        mid = (lo + hi) // 2

        if list[mid] == tgt:
            found = mid
        elif tgt < list[mid]:
            hi = mid - 1
        else:
            lo = mid +1
    return found

#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-#