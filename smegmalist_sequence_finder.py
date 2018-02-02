import urllib.request
import openpyxl
import os

#takes a given gene name and finds its aa sequence from smegmalist
def decode(gene):
    #uses smegmalists's php format to search for genes
    with urllib.request.urlopen("http://mycobrowser.epfl.ch/smegmasearch.php?gene+name=" +
        gene + "&submit=Search") as response:

        #load the search results as an html file, then covert to string
        html = response.read()
        html_string = html.decode("utf-8")

        start = html_string.find("<small>")
        if (start != -1):
            end = html_string.rfind("</small>")

            #pull out the code from the html file
            #7 represents the length of string "<small>"
            code = html_string[start+7:end]


            #removes remaining html artifacts in the code
            code = code.replace("</small><br /><small>", "")
        else:
            code = "none"
            
        return code

#runs the program, looking at each gene name and filling the result
#column with that gene's aa sequence
def runProgram():

    userInput = getInput()

    workbook = openpyxl.load_workbook(userInput["workbook"], data_only = True)
    sheet = workbook.get_sheet_by_name(userInput["sheet"])
    sheet[userInput["results"] + "1"] = "Sequence"

    #Assume a lable row
    row = 2
    maxRow = sheet.max_row

    #loop through the spreadsheet, filling in the aa sequences
    while (row <= maxRow):
        value = sheet[userInput["data"] + str(row)].value
        sheet[userInput["results"] + str(row)] = decode(value)
        row += 1
    workbook.save(userInput["workbook"])
    print("Successfully saved")

#asks the user for the desired excel file, sheet, and columns for data input and sequence output
def getInput():
    print("This program automatically finds gene sequences from smegmalist. All of the following fields are case sensitive. Column letters must be capitalized\n\n\n")
    workbookName = input("excel file name (file must be in same directory as program): ")
    sheet = input("sheet name: ")
    dataColumn = input("column to find gene names (give a letter): ")
    resultColumn = input("column to place generated gene sequences (give a letter): ")
    return {"workbook":workbookName, "sheet":sheet, "data":dataColumn, "results":resultColumn}

runProgram()


