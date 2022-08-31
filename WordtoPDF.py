## Libraries
import win32com.client

def WordtoPDF():
    wdFormatPDF = 17

    ## User inputs
    print("Enter a input file location: ", end = "")
    inputFile = input() + ".docx"

    print("Enter a output file location: ", end = "")
    outputFile = input() + ".pdf"

    ## Convert Word to PDF
    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(inputFile)
    doc.SaveAs(outputFile, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

    print("Conversion completed!")

WordtoPDF()
