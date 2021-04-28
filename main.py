
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





print("cleaning for program.....")

def cleaning():
    
    b = os.listdir("photos/")
    for i in b:
        os.remove("photos/"+i)

    c = os.listdir("cutedtable/")
    for i in c:
        os.remove("cutedtable/"+i)

cleaning()
print("cleaning completed")


def pdftoimage():
    
    evenpdflist = os.listdir("pdfs/")
    k = 0
    for j in evenpdflist:


        images = convert_from_path('pdfs/'+j)
        
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


def replace_trash(unicode_string):
     for i in range(0, len(unicode_string)):
         try:
             unicode_string[i].encode("ascii")
         except:
              #means it's non-ASCII
              unicode_string=unicode_string[i].replace(" ") #replacing it with a single space
     return unicode_string



print("program completed")
