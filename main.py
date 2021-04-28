
from PyPDF2 import PdfFileWriter, PdfFileReader
import PyPDF2
from pdf2image import convert_from_path
import cv2
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
try:
    from PIL import Image
except ImportError:
    import Image
import pytesseract
import os
from skimage import io
import crop_rows
import xlwt
from xlwt import Workbook





print("cleaning for program.....")

def cleaning():

    a = os.listdir("splitedpdf/")
    for i in a:
        os.remove("splitedpdf/"+i)
    
    b = os.listdir("photos/")
    for i in b:
        os.remove("photos/"+i)

    c = os.listdir("cutedtable/")
    for i in c:
        os.remove("cutedtable/"+i)

    d = os.listdir("columns/")
    for i in d:
        os.remove("columns/"+i)

    e = os.listdir("cells/")
    for i in e:
        os.remove("cells/"+i)

    f = os.listdir("excel/")
    for i in f:
        os.remove("excel/"+i)

cleaning()
print("cleaning completed")


def splitEvenPageInPdf():

    pdflist = os.listdir("pdfs/")
    j = 0
    for i in pdflist:
        #print(i)
        pdf_file = "pdfs/" + i

        pdf = PdfFileReader(pdf_file)

        output_filename_even = "splitedpdf/pdf" + str(j) + ".pdf"
        j = j + 1

        pdf_writer_even = PdfFileWriter()

        for page in range(pdf.getNumPages()):
            current_page = pdf.getPage(page)
            if page % 2 == 0:
                pass
            else:
                pdf_writer_even.addPage(current_page)

        with open(output_filename_even, "wb") as out:
            pdf_writer_even.write(out)
            #print("created", output_filename_even)


splitEvenPageInPdf()
print("pdf split opration completed")



def splitsizepdf():
    evenpdflist = os.listdir("splitedpdf/")

    for pdf in evenpdflist:

        input_pdf = PdfFileReader("splitedpdf/"+pdf)
        pagenum = input_pdf.getNumPages()
        stepname = pdf
        stepname = stepname.replace(".pdf","")
        pagecount = 0
        if(input_pdf.getNumPages() > 15):
            j=0
            output = PdfFileWriter()
            for i in range(pagenum):

                output.addPage(input_pdf.getPage(i))
                pagecount += 1

                if(pagecount == 10):
                    output_filename_even = "splitedpdf/"+ stepname + "_" + str(j) + ".pdf"
                    j = j + 1
                    pagecount = 0
                    with open(output_filename_even, "wb") as output_stream:
                        output.write(output_stream)
                    
                    output = PdfFileWriter()

            output_filename_even = "splitedpdf/"+ stepname + "_" + str(j) + ".pdf"
            j = j + 1
            with open(output_filename_even, "wb") as output_stream:
                        output.write(output_stream)


            os.remove("splitedpdf/"+pdf)



splitsizepdf()
print("pdf size fixing done")


def pdftoimage():
    
    evenpdflist = os.listdir("splitedpdf/")
    k = 0
    for j in evenpdflist:


        images = convert_from_path('splitedpdf/'+j)
        
        for i in range(len(images)):
        
            # Save pages as images in the pdf
            images[i].save('photos/page'+ str(k) +'_'+ str(i) +'.jpg', 'JPEG')

        k = k + 1    



pdftoimage()
print("images converted")

photoslist = os.listdir("photos/")
j=0

for p in photoslist:

    #print(p)
    # Read input image
    image = "photos/" + p

    img = cv2.imread(image, cv2.IMREAD_COLOR)

    # Convert to gray scale image
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)

    # Simple threshold
    _, thr = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY)

    # Morphological closing to improve mask
    close = cv2.morphologyEx(255 - thr, cv2.MORPH_CLOSE, cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (3, 3)))

    # Find only outer contours
    contours, _ = cv2.findContours(close, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)

    # Save images for large enough contours
    areaThr = 3000
    i = 0
    for cnt in contours:
        area = cv2.contourArea(cnt)
        if (area > areaThr):
            i = i + 1
            x, y, width, height = cv2.boundingRect(cnt)
            cv2.imwrite('cutedtable/table' + str(j) + '.png', img[y:y+height-1, x:x+width-1])
            j = j + 1

print("table saparation done")


#########################################################

def HansHirseAlgo(f,n):

    img = cv2.cvtColor(io.imread('cutedtable/'+ f ), cv2.COLOR_RGB2BGR)

    img = cv2.resize(img,(1421,580))

    # Inverse binary threshold grayscale version of image
    img_thr = cv2.threshold(cv2.cvtColor(img, cv2.COLOR_BGR2GRAY), 128, 255, cv2.THRESH_BINARY_INV)[1]

    # Count pixels along the y-axis, find peaks
    thr_y = 200
    y_sum = np.count_nonzero(img_thr, axis=0)
    peaks = np.where(y_sum > thr_y)[0]

    # Clean peaks
    thr_x = 50
    temp = np.diff(peaks).squeeze()
    idx = np.where(temp < thr_x)[0]
    peaks = np.concatenate(([0], peaks[idx+1]), axis=0) + 1

    # Save sub-images
    for i in np.arange(peaks.shape[0] - 1):
        if(i==4 or i==5 or i==6 or i==7 or i==8):
            cv2.imwrite('columns/column' + str(n) + '_' + str(i) + '.png', img[:, peaks[i]:peaks[i+1]])

#########################################################

def filteringColumn():

    cloumnlist = os.listdir("columns/")

    for i in cloumnlist:
        #print(i)
        img = cv2.imread("columns/"+i)
        height , width , _ = img.shape

        if(width < 42):
            os.remove("columns/"+i)

        elif(width>50 and width <400):
            os.remove("columns/"+i)




#########################################################

tablelist = os.listdir("cutedtable/")
col = 0
for t in tablelist:

    HansHirseAlgo(t , col)
    col = col + 1

print("column abstraction completed")

filteringColumn()
print("filtering column completed")


celllist = os.listdir("columns/")
cell = 0
for cellfile in celllist:
    crop_rows.box_extraction("columns/"+cellfile, "cells/",cell)
    cell = cell + 1

print("cell abstraction completed")



def filteringcells():

    cloumnlist = os.listdir("cells/")

    for i in cloumnlist:
        #print(i)
        img = cv2.imread("cells/"+i)
        height , width , _ = img.shape

        if(width < 45 and height < 35):
            os.remove("cells/"+i)




filteringcells()
print("filtering cell completed")


def replace_trash(unicode_string):
     for i in range(0, len(unicode_string)):
         try:
             unicode_string[i].encode("ascii")
         except:
              #means it's non-ASCII
              unicode_string=unicode_string[i].replace(" ") #replacing it with a single space
     return unicode_string

def creatingexcel():

    # Workbook is created
    wb = Workbook()
    
    # add_sheet is used to create sheet.
    sheet = wb.add_sheet('amazon_receipt')

    textlist = os.listdir("cells/")
    sheet.write(0,0,"Discription")
    sheet.write(0,1,"Quantity")
    dis = 1
    qty = 1
    for i in textlist:
        #print(i)
        img = cv2.imread("cells/"+i)
        img = cv2.threshold(img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
        width , height , _ = img.shape

        if(width > 450):
            
            text = pytesseract.image_to_string(Image.open("cells/"+i), config='--oem 3 --psm 6', lang='eng')
            #text = replace_trash(text)
            print(i)
            print(text)
            sheet.write(dis,0,text)
            dis += 1
        
        else:
            
            text = pytesseract.image_to_string(Image.open("cells/"+i), config='--oem 3 --psm 6', lang='eng')
            #text = replace_trash(text)
            print(i)
            print(text)
            sheet.write(qty,1,text)
            qty += 1

    wb.save('excel/data.xls')

#creatingexcel()
print("excel created")


print("program completed")