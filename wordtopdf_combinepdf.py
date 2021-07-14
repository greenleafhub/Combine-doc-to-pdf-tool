#pip install docx2pdf, pypiwin32, PyPDF2

#>>>Doc to docx

import glob
import win32com.client
import os

word = win32com.client.Dispatch("Word.Application")
word.visible = 0

for i, doc in enumerate(glob.iglob("*.doc")):
    in_file = os.path.abspath(doc)
    wb = word.Documents.Open(in_file)
    docname = doc.rstrip('.doc')
    print(docname)
    out_file = os.path.abspath("{}.docx".format(docname))
    print(out_file)
    wb.SaveAs2(out_file, FileFormat=16) # file format for docx
    wb.Close()

word.Quit()



#>>>Docx to pdf

from docx2pdf import convert
import os


# Get all the word filenames.
docxList = []
for filename in os.listdir('.'):
    if filename.endswith('.docx') :
        docxList.append(filename)

for file in docxList:
    print(file)
    convert(file)
    


#>>>combine pdf

   #! python3
   # combinepdf.py - Combines all the PDFs in the current working directory into
   # into a single PDF.



import PyPDF2, os
from ctypes import wintypes, windll
from functools import cmp_to_key

#sort like windows: https://stackoverflow.com/questions/4813061/non-alphanumeric-list-order-from-os-listdir/48030307#48030307
def winsort(data):
    _StrCmpLogicalW = windll.Shlwapi.StrCmpLogicalW
    _StrCmpLogicalW.argtypes = [wintypes.LPWSTR, wintypes.LPWSTR]
    _StrCmpLogicalW.restype  = wintypes.INT

    cmp_fnc = lambda psz1, psz2: _StrCmpLogicalW(psz1, psz2)
    return sorted(data, key=cmp_to_key(cmp_fnc))


# Get all the PDF filenames.
pdfFiles = []
for filename in os.listdir('.'):
    if filename.endswith('.pdf'):
        pdfFiles.append(filename)

#replaced pdfFiles.sort(key = str.lower)
pdfFiles = winsort(pdfFiles)
print(pdfFiles)


pdfWriter = PyPDF2.PdfFileWriter()

# Loop through all the PDF files.
for filename in pdfFiles:
    pdfFileObj = open(filename, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
# Create the blank page starting with ‘000_blank’ located in the same directory.
    if filename.startswith('000_blank'):
        blank = pdfReader.getPage(0)
    if not filename.startswith('000_blank'):
        print('Opened ',filename)
# Loop through all the pages and add them.
        for pageNum in range(pdfReader.numPages):
            pageObj = pdfReader.getPage(pageNum)
            pdfWriter.addPage(pageObj)
            print('Combining ',filename, (pageNum +1), '...')

        totalpage = int(pdfReader.numPages)
        if totalpage % 2 == 1:
            pdfWriter.addPage(blank)
    else:
        pass


# Save the resulting PDF to a file.
pdfOutput = open('combined.pdf', 'wb')
pdfWriter.write(pdfOutput)
pdfOutput.close()
print('Combined all to combined.pdf \o/')



#>>>pdf rotate

import PyPDF2

pdf_in = open('combined.pdf', 'rb')
pdf_reader = PyPDF2.PdfFileReader(pdf_in)
pdf_writer = PyPDF2.PdfFileWriter()

numofpages = pdf_reader.numPages

numrotated = 0 
for pagenum in range(numofpages):
    page = pdf_reader.getPage(pagenum)
    mb = page.mediaBox
    if (mb.upperRight[0] > mb.upperRight[1]) and (page.get('/Rotate') is None):
        page.rotateCounterClockwise(90)
        numrotated = numrotated + 1
    pdf_writer.addPage(page)


if (numrotated) == 0:
    print('No rotation needed!')
else:
    pdf_out = open('combined_rotated.pdf', 'wb')
    pdf_writer.write(pdf_out)
    pdf_out.close()
    print (str(numrotated) + " of " + str(numofpages) + " pages were rotated")

pdf_in.close()


#https://stackoverflow.com/questions/48382889/change-orientation-of-landscape-pages-only-in-pdf





