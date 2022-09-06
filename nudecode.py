#importing the required modules
import requests as rq
import easygui
import imageio
import sys
import os
import json
import xlsxwriter
from os import listdir
from pathlib import Path
import tkinter as tk
from tkinter import filedialog
from tkinter import *

#url for the deepai org nsfw module api
url = 'https://api.deepai.org/api/nsfw-detector'

#api key for accessing api online
apiKey = {
   'api-key':'b33fbf9f-7d29-42ec-ae85-98582acf0092'
}

#function for getting path to folder
def findDirectoryOrPath():

    global filepath
    # get a directory path by user
    filepath=filedialog.askdirectory(initialdir=r"C:\Users\User\Pictures", title="Dialog box")
    folderPath=Label(softwareWindow, text=filepath, font=('ariel 14'))
    folderPath.pack(pady=20)
    print(filepath, end='\n')

#function for finding images in a folder
def imageFinder(filepath):

    pathImageList = [ ]
    imageCount = 0
    for image in os.listdir(filepath):
        if(image.endswith(".png") or image.endswith(".jpg") or image.endswith(".jpeg")):
            imageCount+=1
            imageDirectory = filepath + '/' + image
            pathImageList.append(imageDirectory)
            del imageDirectory
        else:
            continue

    print(imageCount, end='\n')

    return pathImageList

# Master list for collecting all image and nude image details as a list
imageDataCollection = [ ]
nudeImageDataCollection = [ ]

#function for classifying images and finding nude images while creating the excel report
def sendImage(pathImageList):

    # variables for counting the total images and the nude images
    imageCount = 0
    nudeCount = 0

    for image in pathImageList:

        imageCount += 1
        imageInformationList = [ ]
        nudeImageInformationList = [ ]
        imageInformationList.append(imageCount)
        splitDirectory = image.split('/')
        imageInformationList.append(splitDirectory[5])
        imageInformationList.append(image)
        imageDetectionInformation = rq.post(url, files = {'image':open(image,'rb')}, headers = apiKey )
        imageDetectionInformationJSON = imageDetectionInformation.json()

        # printing the recieved skin and part analysis report after parsing for reference
        print(imageDetectionInformation.text, end='\n')

        skinExposed = imageDetectionInformationJSON["output"]["nsfw_score"]
        skinExposedPercentage = skinExposed * 100
        imageInformationList.append(skinExposedPercentage)

        # classifying the image based on skin percentage exposed
        if skinExposedPercentage < 45:
            imageInformationList.append('Dressed')
        elif skinExposedPercentage >= 45 and skinExposedPercentage < 85:
            imageInformationList.append('Semi Nude')
        elif skinExposedPercentage >=85:

            # list for inputting human body parts detected
            detectedParts = [ ]
            nudeCount += 1
            imageInformationList.append('Nude')

            # nude image information appended as a list
            nudeImageInformationList.append(nudeCount)
            nudeImageInformationList.append(splitDirectory[5])
            nudeImageInformationList.append(image)
            nudeImageInformationList.append(skinExposedPercentage)

            # gathering information on the detected parts exposed and appending it
            for part in imageDetectionInformationJSON["output"]["detections"]:
                detectedParts.append(part["name"])
            detectedPartCombined = ' '.join(detectedParts)
            nudeImageInformationList.append(detectedPartCombined)

            #appending the collected listed information of nude image to nude data bank
            nudeImageDataCollection.append(nudeImageInformationList)
            del detectedParts
            del detectedPartCombined

        # appending all the images listed information to all image data bank
        imageDataCollection.append(imageInformationList)


# UI for selecting the folder whose path needs to be found
softwareWindow = tk.Tk()
softwareWindow.geometry('500x500')
softwareWindow.title('Adult Content Detector')
softwareWindow.grid_rowconfigure(0, weight = 1)
softwareWindow.grid_columnconfigure(0, weight = 1)
softwareWindow.configure(background='#F08D7E')
title=Label(softwareWindow,background='#EFA18A', font=('Times',20,'bold'))
folderUploadButton = Button(softwareWindow, text='Select Folder', command = findDirectoryOrPath)
folderUploadButton.configure(background='#E2BAB1', foreground='white',font=('Times',20,'bold'))
folderUploadButton.pack(side=TOP,pady=150)
softwareWindow.mainloop()

# calling image finder function to get all the images in the folder
pathImageList = imageFinder(filepath)

# calling function to identify nude images in the folder
sendImage(pathImageList)

# printing the stored data in all image list and nude image list
print(imageDataCollection, end='\n')
print(nudeImageDataCollection, end='\n')

# assigning the total number of images and nude images present
totalImageCount = len(imageDataCollection)
nudeTotalImageCount = len(nudeImageDataCollection)

#Converting list to tuple for ease of creating excel report
imageDataCollectionTuple = tuple(imageDataCollection)
nudeImageDataCollectionTuple = tuple(nudeImageDataCollection)

print("CREATING EXCEL WORKBOOK", end='\n')
# creating the excel workbook for image classification report
workbook = xlsxwriter.Workbook('Image Analysis Report.xlsx')

# creating the excel worksheet for all image classification and analysis report
worksheet1 = workbook.add_worksheet("Image Report")

worksheet1.write('A1','Sl. No.')
worksheet1.write('B1','Image Name')
worksheet1.write('C1','Image Path')
worksheet1.write('D1','Exposed Skin Percentage')
worksheet1.write('E1','Classified Category')

# creating the excel worksheet for nude image report
worksheet2 = workbook.add_worksheet("Nude Image Report")

worksheet2.write('A1','Sl. No.')
worksheet2.write('B1','Image Name')
worksheet2.write('C1','Image Path')
worksheet2.write('D1','Exposed Skin Percentage')
worksheet2.write('E1','Body Part Exposed')

# initialising row and column for excel sheet 1 for all image analysis report
row1 = 1
col1 = 0

# initialising row and column for excel sheet 2 for nude image analysis report
row2 = 1
col2 = 0

# entering all the images analysis information report in the excel sheet
for slno, name, pathName, exposedSkinPercentage, classifiedCategory in (imageDataCollectionTuple):
    worksheet1.write(row1, col1, slno)
    worksheet1.write(row1, col1 + 1, name)
    worksheet1.write(row1, col1 + 2, pathName)
    worksheet1.write(row1, col1 + 3, exposedSkinPercentage)
    worksheet1.write(row1, col1 + 4, classifiedCategory)
    row1 += 1

worksheet1.write(row1 + 1, col1, 'Total Images')
worksheet1.write(row1 + 1, col1 + 1, totalImageCount)

# entering all the images analysis information report in the excel sheet
for slno, name, pathName, exposedSkinPercentage, partExposed in (nudeImageDataCollectionTuple):
    worksheet2.write(row2, col2, slno)
    worksheet2.write(row2, col2 + 1, name)
    worksheet2.write(row2, col2 + 2, pathName)
    worksheet2.write(row2, col2 + 3, exposedSkinPercentage)
    worksheet2.write(row2, col2 + 4, partExposed)
    row2 += 1

worksheet2.write(row2 + 1, col2, 'Total Nude Images')
worksheet2.write(row2 + 1, col2 + 1, nudeTotalImageCount)

workbook.close()
print("WORKBOOK CREATED SUCCESSFULLY !", end='\n')
print("Loading Workbook...")
# opening the excel workbook saved and excel using python
workbookPath = Path('../Project/Image Analysis Report.xlsx').resolve()
os.system(f'start excel.exe "{workbookPath}"')
