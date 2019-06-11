# -*- coding: utf-8 -*-
"""
Created on Thu Mar 28 00:42:09 2019

@author: bhaik
"""

from pyzbar import pyzbar
import cv2
from xlwt import Workbook
from datetime import datetime
from datetime import date

wb = Workbook()
sheet1 = wb.add_sheet("Sheet 1")
sheet1.write(0,0, "S. No.")
sheet1.write(0,1, "Roll. No.")
sheet1.write(0,2, "Date")
sheet1.write(0,3, "In Time")
sheet1.write(0,4, "Book S. No.")
sheet1.write(0,5, "Out_Time")
wb.save("Xlib.xls")

count = 1
lis = []
c=True
d=1

while True:
    
    if len(lis) == 0:
        cnt = []
        today = date.today()
        t = datetime.time(datetime.now())
            
        
        cap = cv2.VideoCapture(0)
        #cap = cv2.VideoCapture(1)
        
        while True:
            ret,frame = cap.read()
            cv2.imshow("frame", frame)
            key = cv2.waitKey(1)
            if key == 27:
                cv2.imwrite("123.jpg", frame)
                break
            
        cap.release()    
        cv2.destroyAllWindows()
        
        
        path = "C:\\Users\\bhaik\\OneDrive\\Desktop\\Vibhav_projects\\Finalised\\Library\\123.jpg"
        
        image = cv2.imread(path)
        
        barcodes = pyzbar.decode(image)
        
        for barcode in barcodes:
            (x, y, w, h) = barcode.rect
            cv2.rectangle(image, (x, y), (x + w, y + h), (0, 0, 255), 2)
            barcodeData = barcode.data.decode("utf-8")
            barcodeType = barcode.type
            text2 = "{} ({})".format(barcodeData, barcodeType)
            cv2.putText(image, text2, (x, y - 10), cv2.FONT_HERSHEY_SIMPLEX,
        		0.5, (0, 0, 255), 2)
            
        cap = cv2.VideoCapture(0)
        
        while True:
            ret, frame = cap.read()
            
            cv2.imshow("frame", frame)
            
            key = cv2.waitKey(1)
            if key == 27:
                cv2.imwrite("123.jpg", frame)
                break
            
        cap.release()    
        cv2.destroyAllWindows()
        
        path = "C:\\Users\\bhaik\\OneDrive\\Desktop\\Vibhav_projects\\Finalised\\Library\\123.jpg"
        
        image = cv2.imread(path)
        
        barcodes = pyzbar.decode(image)
        
        for barcode in barcodes:
            (x, y, w, h) = barcode.rect
            cv2.rectangle(image, (x, y), (x + w, y + h), (0, 0, 255), 2)
            barcodeData = barcode.data.decode("utf-8")
            barcodeType = barcode.type
            text = "{} ({})".format(barcodeData, barcodeType)
            cv2.putText(image, text, (x, y - 10), cv2.FONT_HERSHEY_SIMPLEX,
        		0.5, (0, 0, 255), 2)
            
        sheet1.write(count,0, count)
        sheet1.write(count,1, text2)
        sheet1.write(count,2, str(today))
        sheet1.write(count,3, str(t))
        sheet1.write(count,4, text)
        
        cnt.append(count)
        cnt.append(text)
        cnt.append(text2)
        lis.append(cnt)
        
        print(lis)
        
        print("-------------------------------")
        print("You have entered your book")
        print("Your Roll No:")
        print(text2)
        print("Information in Barcodes:")
        print(text)
        print("-------------------------------")
        cv2.imshow("Image", image)
        wb.save("Xlib.xls")
        cv2.waitKey(0)
        
        cv2.destroyAllWindows()
        
    else:
        
        c=True
        cnt = []
        today = date.today()
        t = datetime.time(datetime.now())
        
        cap = cv2.VideoCapture(0)
        while True:
            ret,frame = cap.read()
            cv2.imshow("frame", frame)
            key = cv2.waitKey(1)
            if key == 27:
                cv2.imwrite("123.jpg", frame)
                break
            
        cap.release()    
        cv2.destroyAllWindows()
        
        path = "C:\\Users\\bhaik\\OneDrive\\Desktop\\Vibhav_projects\\Finalised\\Library\\123.jpg"
        
        image = cv2.imread(path)
        
        barcodes = pyzbar.decode(image)
        
        for barcode in barcodes:
            (x, y, w, h) = barcode.rect
            cv2.rectangle(image, (x, y), (x + w, y + h), (0, 0, 255), 2)
            barcodeData = barcode.data.decode("utf-8")
            barcodeType = barcode.type
            text2 = "{} ({})".format(barcodeData, barcodeType)
            cv2.putText(image, text2, (x, y - 10), cv2.FONT_HERSHEY_SIMPLEX,
        		0.5, (0, 0, 255), 2)
        
        cap = cv2.VideoCapture(0)
        
        while True:
            ret, frame = cap.read()
            
            cv2.imshow("frame", frame)
            
            key = cv2.waitKey(1)
            if key == 27:
                cv2.imwrite("123.jpg", frame)
                break
            
        cap.release()    
        cv2.destroyAllWindows()
        
        path = "C:\\Users\\bhaik\\OneDrive\\Desktop\\Vibhav_projects\\Finalised\\Library\\123.jpg"
        
        image = cv2.imread(path)
        
        barcodes = pyzbar.decode(image)
        
        for barcode in barcodes:
            (x, y, w, h) = barcode.rect
            cv2.rectangle(image, (x, y), (x + w, y + h), (0, 0, 255), 2)
            barcodeData = barcode.data.decode("utf-8")
            barcodeType = barcode.type
            text = "{} ({})".format(barcodeData, barcodeType)
            cv2.putText(image, text, (x, y - 10), cv2.FONT_HERSHEY_SIMPLEX,0.5, (0, 0, 255), 2)
            
            
        for l in lis:
            
            if text == l[1]:
                c = False
                remove1 = l

        if c == True:
            cnt.append(count)
            cnt.append(text)
            cnt.append(text2)
            lis.append(cnt)
            sheet1.write(count,0, count)
            sheet1.write(count,1, text2)
            sheet1.write(count,2, str(today))
            sheet1.write(count,3, str(t))
            sheet1.write(count,4, text)
            print("-------------------------------")
            print("you have entered your book")
            print("Roll No")
            print(text2)
            print("Information in Barcodes:")
            print(text)
            print("you have entered your book")
            print("-------------------------------")
            wb.save("Xlib.xls")
            print(lis)

        if c == False:
            t1 = datetime.time(datetime.now())
            sheet1.write(remove1[0],5, str(t1))
            wb.save("Xlib.xls")
            print("your_entry_closed_sucessfully")
            lis.remove(remove1)
            count = count - 1
            print(remove1)
            print(lis)
            
    count = count + 1
