#Module for first part
import urllib.request
# Modules for second part
#Library that helps with file handling for zip files
import zipfile
#Library that helps to interact with operating system
import os
#Module for the third part
#module csv provides a handler for csv files
import csv
#Module for fifth part
#Module that will help to open and create excel files
import xlsxwriter

######################################################
# Step 1: Download zip file
######################################################

#initialize variable with the URL to download
#Website: National Stock Exchange of India
#urlOfFileName = "http://www.nseindia.com/content/historical/EQUITIES/2020/MAR/cm26MAR2020bhav.csv.zip"
urlOfFileName = "https://archives.nseindia.com/content/historical/EQUITIES/2020/MAR/cm26MAR2020bhav.csv.zip"

#initialize a variable with the local file in which to store the file
#This is a path on the local desktop
localZipFilePath = "C:/Users/Katja/PythonCodingTraining/DoingStuffwithFiles/cm26MAR2020bhav.csv.zip"

#The following code is a boilerplate to deal with the fact that the website blocks bots/automated scripts
hdr = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36',
    #'Accept': 'text/html, application/xhtml+xml,application/xml;q=0.9,*/*q=0.8',
    #'Accept-Charset':'ISO-8859-1;utf-8,q=0.7,*;q=0.3',
    #'Accept-Encoding':'none',
    #'Accept-Language':'en-US,en;q=0.8',
    #'Connection':'keep-alive'
    }

#The following code will actually download and store the file

#Make the web request by using a web brower
webRequest = urllib.request.Request(urlOfFileName, headers = hdr)

try:
    #Making download request
    page = urllib.request.urlopen(webRequest)
    #save the content of this opage into a variable called 'content'
    #By reading the content of the downloaded page we put the content into the variable 'content'
    content = page.read()
    #Save content to local disk
    #Use open command to create a new file
    output = open(localZipFilePath, 'wb')
        #'w'indicates that we intend to write
        #'b' indicats that this is a binary file, i.e. not a text file
    #Writing content to the file
    #bytearray() returns an array of bytes of the size of content
    #The bytearray can then be written to file
    output.write(bytearray(content))
    #File needs to be closed before moving on
    output.close()
except urllib.error.HTTPError as httpe:
    print(httpe.fp.read())
    print('Looks like to file download did not go through for url = ' + urlOfFileName)

###########################################
# Step 2: Unzip the download file and extract
###########################################

#initiliaze a variable with the local directory in which to extract
#the zip file above
localExtractFilePath = "C:/Users/Katja/PythonCodingTraining/DoingStuffwithFiles/"

#Check if the zip file was downloaded successfully
if os.path.exists(localZipFilePath):
    print("Cool! " + localZipFilePath + ' exists...proceeding')

    # We do not know how many files are in the zip file
    #Hence, initialize array in which to save the names of the files
    listOfFiles = []
    #Open the zip file
    fh = open(localZipFilePath, 'rb')
        # r indicates that file is open in read mode
        # b indicates that this is a binary file

    #binary files require handlers (unlike text files)
    #the zipfile handler knows how to read the list of zipped up files
    zipFileHandler = zipfile.ZipFile(fh)

    #iterating over list in a for-loop
    for filename in zipFileHandler.namelist():
        #Files are extracted
        zipFileHandler.extract(filename, localExtractFilePath)
        #Add to the list of files we have extracted
        listOfFiles.append(localExtractFilePath + filename)
        print("Extracted " + filename + " from the zip file to " + (localExtractFilePath + filename))

    print("In total, we extracted ", str(len(listOfFiles)), ' files.')

    #Close the file
    fh.close()

###################################################################
# Step 3: Parse the extracted CSV and create a table of stock info
###################################################################

#We know in this case that the zip file has only one file which we access by using the index operator
oneFileName = listOfFiles[0]
lineNum = 0
#We will store whatever we read from the csv file in variable 'list of lists'
#More specifically, it will take care of three columns that we care about.
#ticker, percentage change and value traded
listOfLists = []

#with is a construct that allows us to open a file and do stuff with it,
#and not worrying about explicitly opening and closing the file
#When the with block opens, the file is explicitly opended
#When the with block ends, the file is implictly closed
with open(oneFileName,'r') as csvfile:
    #Use of handler csv.reader which reads one line at a time
    #Information required by handler a)delimiter b)how cell values that contain the separator will be special-cased
    lineReader = csv.reader(csvfile, delimiter = ',', quotechar="\"")

    #Actually reading the lines in csv file one by one
    for row in lineReader:
        #keeping track of the line we are by using the variable lineNum
        lineNum = lineNum + 1
        if lineNum == 1:
            print('Skipping the header row')
            continue
        #We know from the header that:
        #column 1 (index = 0) is the stock ticker
        #column 6 (index = 5) is the last closing price
        #column 8 (index = 7) has yesterday's closing price
        #column 10 (index = 9) has the traded value in India Rupees
        symbol = row[0]
        close = row[5]
        prevClose = row[7]
        tradedQty = row[9]
        pctChange = float(close)/float(prevClose) - 1

        #Creating list which contains the three columns we actually care about
        oneResultRow = [symbol, pctChange, float(tradedQty)]

        #Appending oneResultRow to listOfLists to build a running result
        listOfLists.append(oneResultRow)
        print(symbol, "{:,.1f}".format(float(tradedQty)/1e6) + "M INR", "{:,.1f}".format(pctChange*100)+"%")
    #out of for loop
print("Done iterating over the file contents - the file is closed now!")
print("We have stock info for " + str(len(listOfLists)) + " stocks")

##################################################
# Step 4: Sort the list of lists!
##################################################

#lambda function are now used which si 'idiomatic' python

#Use of lambda function to sort list of lists
# Here we sort the list of lists by column 3 (index = 2). The reverse = True means that sort will be descending
listofListsSortedByQty = sorted(listOfLists, key = lambda x: x[2], reverse = True)
listOfListsSortedByPctChange = sorted(listOfLists, key=lambda x: x[1], reverse=True)

###########################################################
# Step 5: Create excel file with top five moving stocks
###########################################################

#Initialize a variable with the name of the excel file
#The variable should be a path on your local desktop
excelFileName =  "C:/Users/Katja/PythonCodingTraining/DoingStuffwithFiles/SpreadsheetbyPython.xlsx"

#New workbook is created
workbook = xlsxwriter.Workbook(excelFileName)
#Adding worksheet called 'summary' to excel workbook
worksheet = workbook.add_worksheet('Summary')

#Write into worksheet cell by cell
worksheet.write_row('A1',['Top Traded Stocks'] )
#The way to write stuff into the excel is by specifying
#a) the cell address in the standard excel format
#b) the list of values to be written, 1 per cell starting from that address
worksheet.write_row('A2',['Stock', 'PctChange', 'Value Traded (INR)'])

#Writing the 5 most heavily traded stocks by using a for loop
#To make sure that only 5 first numbers are written into file, we use the range function
#e.g. range(5), i.e. python shorthand for a list of the first 5 numbers
for rowNum in range(5):
    #get the corresponding row from our sorted result
    oneRowToWrite = listOfListsSortedByPctChange[rowNum]
    #write this out to the excel spreadsheet
    #Add one to the row number because excel starts indexing at one
    worksheet.write_row('A'+str(rowNum + 3), oneRowToWrite)
workbook.close()