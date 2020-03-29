#Module for first part
import urllib.request
# Modules for second part
#Library that helps with file handling for zip files
import zipfile
#Library that helps to interact with operating system
import os

#initialize variable with the URL to download
#Website: National Stock Exchange of India
urlOfFileName = "http://www.nseindia.com/content/historical/EQUITIES/2020/MAR/cm26MAR2020bhav.csv.zip"

#initialize a variable with the local file in which
#to store the URL
#This is a path on the local desktop
localZipFilePath = "C:/Users/Katja/PythonCodingTraining/DoingStuffwithFiles/cm26MAR2020bhav.csv.zip"

#The following code is a boilerplate to deal with the fact
#that the website blocks bots/automated scripts
hdr = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36',
    'Accept': 'text/html, application/xhtml+xml,application/xml;q=0.9,*/*q=0.8',
    'Accept-Charset':'ISO-8859-1;utf-8,q=0.7,*;q=0.3',
    'Accept-Encoding':'none',
    'Accept-Language':'en-US,en;q=0.8',
    'Connection':'keep-alive'
    }

#The following code will actually downlaod and store the file

#Make the web request by using a web brower
webRequest = urllib.request(urlOfFileName, header: hdr)

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
        #b indicates that this is a binary file

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