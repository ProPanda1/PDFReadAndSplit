from __future__ import division, print_function, absolute_import, unicode_literals
import subprocess
import sys
import os
import tkinter
import fitz
from tkinter import filedialog
import pathlib

root = tkinter.Tk()
adobePathCreated = False

file = pathlib.Path(__file__).parent
print(str(file) + r"\AdobeLocation.txt")
if not os.path.exists(str(file) + r"\AdobeLocation.txt"):
  fileObject = open("AdobeLocation.txt","w+")
  AdobeLocation = filedialog.askopenfile(parent=root,mode='rb',title='Choose the Adobe Reader Executable')
  AdobeLocationName = str(AdobeLocation.name)
  fileObject.write(AdobeLocationName)
  fileObject.close
  adobePathCreated = True
print(os.stat(str(file) + r"\AdobeLocation.txt"))
if adobePathCreated == True:
  fileObject = open("AdobeLocation.txt","r+")
  AdobeLocationName = fileObject.read()
  fileObject.close
elif os.path.exists(str(file) + r"\AdobeLocation.txt"):
  fileObject = open("AdobeLocation.txt","r+")
  AdobeLocationName = fileObject.read()
  fileObject.close
else:
  fileObject = open("AdobeLocation.txt","w+")
  AdobeLocation = filedialog.askopenfile(parent=root,mode='rb',title='Choose the Adobe Reader Executable')
  AdobeLocationName = str(AdobeLocation.name)
  fileObject.write(AdobeLocationName)
  fileObject.close


fileObject = open("PrinterLocation.txt","r")
printerLocationName = fileObject.read()
fileObject.close


PDFfile = filedialog.askopenfile(parent=root,mode='rb',title='Choose the PDF')


print(str())
print(str(PDFfile.name))

with fitz.open(PDFfile) as doc:
    text = ""
    documentCounter = 0
    pages = [0]
    for page in doc:
      if documentCounter != 0:
        if "Work Order Number:" in page.getText():
          doc2 = fitz.open()
          documentCounter += 1
          firstpage = pages[-1]
          print(firstpage)
          doc2.insertPDF(doc, from_page = firstpage, to_page = page.number - 1, start_at = 0)
          documentName = "WorkOrderPrintNumber" + str(documentCounter - 1) + ".pdf"
          print(documentName)
          doc2.save(documentName)
          pages.append(page.number)
        elif page.number == doc.pageCount - 1:
          doc2 = fitz.open()
          documentCounter += 1
          firstpage = pages[-1]
          print(firstpage)
          doc2.insertPDF(doc, from_page = firstpage, to_page = page.number - 1, start_at = 0)
          documentName = "WorkOrderPrintNumber" + str(documentCounter - 1) + ".pdf"
          print(documentName)
          doc2.save(documentName)
      else:
        documentCounter = 1
documentCounter += -1
print(pages)
print(documentCounter)

print(str(file))
# ref: https://geonet.esri.com/thread/59446
# ref: https://helpx.adobe.com/jp/acrobat/kb/510705.html


# acroread = r'C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe'
acrobat = AdobeLocationName

# '"%s"'is to wrap double quotes around paths
# as subprocess will use list2cmdline internally if we pass it a list
# which escapes double quotes and Adobe Reader doesn't like that

for x in range(1, documentCounter):
    print(str(file) + str(r"\WorkOrderPrintNumber" + str(x) + ".pdf"))
    cmd = acrobat + r" /N " + r" /T " + str(file) + str(r"\WorkOrderPrintNumber" + str(x) + ".pdf ") + printerLocationName

    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    stdout, stderr = proc.communicate()
    exit_code = proc.wait()


for x in range(1, documentCounter+1):
    os.remove(str(file) + str(r"\WorkOrderPrintNumber" + str(x) + ".pdf "))


