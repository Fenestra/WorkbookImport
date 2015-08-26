This project can be run by copying the contents of the "dist" directory to someplace and renaming it something useful (like "jsonTesting").


You can then run this from the command line by running the following command:

workbookImport sourceFileName destinationDirectory

examples:

workbookImport c:\data\MA-10000L.xls  

this line will put the output MA-10000L.json file in the same directory that the batch file is in



workbookImport sampleSpreadsheet.xlsx c:\data

this line will assume the samplespreadsheet.xlsx file is in the same directory as the batch file and the output will be placed in the c:\data directory