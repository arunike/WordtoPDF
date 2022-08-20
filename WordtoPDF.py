import os
import win32com.client

wdFormatPDF = 17

inputFile = os.path.abspath(r"") #Inputfile absolute location path
outputFile = os.path.abspath(r"") #Outputfile absolute location path
word = win32com.client.Dispatch('Word.Application')
doc = word.Documents.Open(inputFile)
doc.SaveAs(outputFile, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()
