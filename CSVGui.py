import tkinter as tk
from tkinter import filedialog
from tkinter import *
from tkinter import ttk
from tkinter import StringVar
import time
from subprocess import Popen
import csv
import pyautogui

######################################

currentFile = None
running = False
readyToRock = False
counter = 0
barcodesLength = 0
progressBarValue = 0
barcodeSet = []


######################################

class PartcodeClass(object):
    def __init__(self, partcode,description,quantity):
        self.partcode = partcode
        self.description = description
        self.quantity = quantity

def createBarcodeList(csvfile):
    global barcodesLength,barcodeSet
    rowNum = 0
    for row in csvfile:
        if rowNum == 0:
            tags = row
        else:
            if len(str(row[3])) > 2:
                tempQuantity = '1'
            else:
                tempQuantity = row[3]
            barcodeSet.append(PartcodeClass(row[0],row[2],tempQuantity))
        rowNum +=1
    barcodesLength = len(barcodeSet)
    print('Amount of barcodes: ' + str(barcodesLength))
    pb["maximum"] = barcodesLength

######################################

def enterItem():    
    #Launches enter item dialog
    print('\nPressing: i')
    pyautogui.press('i')

def createPartcode(partcode, description, quantity):
    
    print('\n--------- New Item ---------')
    
    #Shifts back to Item field
    print('\n Pressing: Shift + Tab')
    pyautogui.hotkey('shift', 'tab')
    
    #Enters partcode
    print('  > Partcode: ' + partcode)
    pyautogui.typewrite(partcode)
    
    #Shifts forward to description field
    print(' Pressing: Tab')
    pyautogui.press('tab')
    
    #Enters description field
    print('  > Description: ' + description)
    pyautogui.typewrite(description)

    #Shifts forward to quantity field
    print(' Pressing: Tab')
    pyautogui.press('tab')

    #Enters quantity
    print('  > Quantity: ' + quantity)
    pyautogui.typewrite(quantity)
    
    #Shifts back to OK Button
    print(' Pressing: Shift Tab x4')
    for a in range (0,4):
        print('  > Tab ' + str(a+1))
        pyautogui.hotkey('shift', 'tab')
        
    #Shifts to OK Button
    print(' Pressing: Enter')
    pyautogui.press('enter')

######################################

def loadCSV():
    global barcodeSet,currentFile, barcodesLength,readyToRock
    
    barcodeSet = []    
    currentFile = tk.filedialog.askopenfilename(filetypes=(('CSV Files','.csv'),('All files', '*')))
    print(currentFile)
    
    if not currentFile:
        print('No File Found')
        csvFileName.set('Load CSV File')
        barcodeListText.set('0')
        firstBarcode.set('N/A')
        readyToRock = False
    else:
        print('Found File')
        tempFileName = currentFile.split("/")
        tempFileName = tempFileName[len(tempFileName)-1]
        csvFileName.set(tempFileName)
        csvData = csv.reader(open(currentFile))
        createBarcodeList(csvData)
        print('>>>>>> Loaded file: ' + tempFileName)
        barcodeListText.set(str(barcodesLength))
        firstBarcode.set(barcodeSet[0].partcode)
        readyToRock = True

def inputParts():
    if running:
        if readyToRock:
            global counter,progressBarValue,barcodesLength
            if counter < barcodesLength:
                enterItem()
                createPartcode(barcodeSet[counter].partcode, barcodeSet[counter].description,barcodeSet[counter].quantity)
                progressBarValue += 1
                pb["value"] = progressBarValue
                counter += 1
            else:
                actionText.set('Completed')
                root.destroy()
            
    root.after(1, inputParts)

def startStop():
    if readyToRock:
        global running
        if (actionText.get() == 'Start'):
            #findClarity()
            timer()
            running = True
            actionText.set('Stop')
        elif (actionText.get() == 'Stop'):
            running = False
            actionText.set('Start')
        else:
            running = False
            actionText.set('Completed')
            root.destroy()

def timer():
    print('Test script runs in:')
    timer = 5
    for a in range (0,timer):
        actionText.set(str(timer-a))
        print(timer-a)
        time.sleep(1);

def findClarity():
    if (pyautogui.locateOnScreen('Images/wizard.png')) != None: 
        clarityWizard = pyautogui.locateOnScreen('Images/wizard.png')
        clarityWizardX,clarityWizardY = pyautogui.center(clarityWizard)
        pyautogui.click(clarityWizardX, clarityWizardY)
        print('Found Clarity Wizard')
    else:
        print('Clarity Wizard not found, requires user input')

######################################

root = Tk()
root.title('CSV to Clarity Partcodes')
root.minsize(250,5)

mainframe = ttk.Frame(root, padding="10 10 10 10")
mainframe.pack()

csvFileName = StringVar()
csvFileName.set('Load CSV File')
actionText = StringVar()
actionText.set('Start')
barcodeListText = StringVar()
barcodeListText.set('0')
firstBarcode = StringVar()
firstBarcode.set('N/A')


loadButton = ttk.Button(mainframe, text='Open CSV', command = loadCSV).pack(ipadx = 5,
                                                                            ipady = 5,
                                                                            padx = 5,
                                                                            pady = 5,
                                                                            side="top")

fileSpecsFrame = ttk.Frame(mainframe, borderwidth=10, relief="solid")
fileSpecsFrame.pack()

static1 = ttk.Label(fileSpecsFrame, text='CSV File: ').grid(column=1, row=1, sticky=(W))
static2 = ttk.Label(fileSpecsFrame, text='No. of Barcodes: ').grid(column=1, row=2, sticky=(W))
static3 = ttk.Label(fileSpecsFrame, text='First Barcode: ').grid(column=1, row=3, sticky=(W))
loadedFileLabel = ttk.Label(fileSpecsFrame, textvariable=csvFileName).grid(column=2, row=1, sticky=(E))
barcodeListLabel = ttk.Label(fileSpecsFrame, textvariable=barcodeListText).grid(column=2, row=2, sticky=(E))
firstBarcodeLabel = ttk.Label(fileSpecsFrame, textvariable=firstBarcode).grid(column=2, row=3, sticky=(E))

pb = ttk.Progressbar(mainframe, orient=HORIZONTAL,
                     length=200, mode='determinate',
                     maximum = barcodesLength,
                     value = progressBarValue)
pb.pack(ipadx = 5, ipady = 5,padx = 5, pady = 5)


bottomFrame = ttk.Frame(mainframe, padding="5 5 5 5")
bottomFrame.pack()

startStopButton = ttk.Button(bottomFrame,
                             textvariable= actionText,
                             command = startStop)
startStopButton.pack(ipadx = 5, ipady = 5,padx = 5, pady = 5,side="left")
quitButton = ttk.Button(bottomFrame, text='Quit', command = root.destroy).pack(ipadx = 5,
                                                                               ipady = 5,
                                                                               padx = 5,
                                                                               pady = 5,
                                                                               side="right")

######################################

root.after(1, inputParts)
root.mainloop()

######################################
