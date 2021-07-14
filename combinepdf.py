 #! python3
   # combinepdf.py - Combines all the PDFs in the current working directory into
   # into a single PDF.

import PyPDF2, os

# Get all the PDF filenames.
pdfFiles = []
for filename in os.listdir('.'):
    if filename.endswith('.pdf'):
        pdfFiles.append(filename)

pdfFiles.sort(key = str.lower)
pdfWriter = PyPDF2.PdfFileWriter()

# Loop through all the PDF files.
for filename in pdfFiles:
    pdfFileObj = open(filename, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
# Create the blank page located in the same directory.
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
pdfOutput = open('combinied.pdf', 'wb')
pdfWriter.write(pdfOutput)
pdfOutput.close()
print('Combined all to allminutes.pdf \o/')
