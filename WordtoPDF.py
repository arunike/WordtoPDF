import os
import win32com.client

wdFormatPDF = 17

inputFile = os.path.abspath(r"C:\Users\r6311\Downloads\College and Career\Resume\Richie Zhou's Resume 2022.docx")
outputFile = os.path.abspath(r"C:\Users\r6311\Downloads\College and Career\Resume\Richie Zhou's Resume 2022.pdf")
word = win32com.client.Dispatch('Word.Application')
doc = word.Documents.Open(inputFile)
doc.SaveAs(outputFile, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()