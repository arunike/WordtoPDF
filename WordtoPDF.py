## Libraries
import win32com.client

wdFormatPDF = 17

## User inputs
print("Enter a input file location: ")
inputFile = input()

print("Enter a output file location: ")
outputFile = input()

## Convert Word to PDF
word = win32com.client.Dispatch('Word.Application')
doc = word.Documents.Open(inputFile)
doc.SaveAs(outputFile, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()

print("Conversion completed!")
