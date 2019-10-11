# Button to preview created barcodes to call 'treeview' widget

import tkinter as tk
from tkinter import filedialog
from tkinter import *
from tkinter import ttk
from tkinter import StringVar
import time
from subprocess import Popen
import csv

######################## INITIAL CONSTANTS ############################

initialDirectory = "/"
print('CSVBarcodes.csv | ' + (time.strftime('%c')))
barcodes = []
categories = ['PARTCODE', 'DESCRIPTION', 'LOCATION', 'QUANTITY', 'PRODUCTION']
barcodes.append(categories)

############################ FUNCTIONS ################################

def addNewBarcodes():

    if var.get() is 1:
        production = 'Print'
    else:
        production = 'Cut'
    
    rangeValue = int(rangeText.get())
    if rangeValue == 0:
        pass
    
    elif rangeValue == 1:
        
        partcode = partcodeText.get()
        description = descriptionText.get()
        location = locationText.get()
        quantity = quantityText.get()

        print(partcode
             + ' | ' + description
             + ' | ' + location
             + ' | ' + quantity
             + ' | ' + production
             )
        
        barcodes.append([partcode,description,location,quantity,production])
        
    else:
        for a in range(0,rangeValue):
            counter = str(a+1)
            
            if (a+1) < 10:
                counter = '0' + str(a+1)
                
            partcode = partcodeText.get()+counter
            description = descriptionText.get()
            location = locationText.get()+ ' ' + str(a+1)
            quantity = quantityText.get()
            
            print(partcode
                  + ' | ' + description
                  + ' | ' + location
                  + ' | ' + quantity
                  + ' | ' + production
                  )

            barcodes.append([partcode,description,location,quantity,production])
    

def getFilenamePath(filename):
    
    #print(filename)
    tempDirectory = filename.split('/')
    #print(tempDirectory)
    currentDirectory = ''
    
    for a in range(0,(len(tempDirectory)-1)):
        if tempDirectory[a] == '':
            currentDirectory += '/'
        else:
            currentDirectory += tempDirectory[a] + '/'
    #print(currentDirectory)
    return currentDirectory


def saveCSV():

    filename = tk.filedialog.asksaveasfilename(initialdir = initialDirectory,
                                              title = "Select file",
                                              filetypes = (("CSV files","*.csv"),("All files","*.*")))
    
    csv.register_dialect('myDialect', lineterminator = '\r')
    
    try:
        with open((filename+'.csv'), 'w') as f:
            writer = csv.writer(f, dialect='myDialect')
            barcodesLength = int(len(barcodes))
            for a in range(0,barcodesLength):
                writer.writerow(barcodes[a])
        f.close()
    except PermissionError:
        warningMsg()

def loadCSV():
    global initialDirectory
    filename = tk.filedialog.askopenfilename(filetypes=(('CSV Files','.csv'),('All files', '*')))
    if filename != '':
        #csvData = csv.reader(open(filename,newline=''),delimiter=',')
        csvData = csv.reader(open(filename))
        initialDirectory = getFilenamePath(filename)
    else:
        initialDirectory = '/'
        
    for row in csvData:
        if row == categories:
            pass
        else:
            barcodes.append(row)


def warningMsg():
    print('Access Denied')
    warrningWindow= Tk()
    warningInfo = ttk.Label(warrningWindow,text = 'File might be open.\nChange name or choose another file.').pack(ipadx = 5,
                                                ipady = 5,
                                                padx = 10,
                                                pady = 10,
                                                side="top")
    closeButton = ttk.Button(warrningWindow,
                        text='Close',
                        command = warrningWindow.destroy).pack(ipadx = 5,
                                                ipady = 5,
                                                padx = 10,
                                                pady = 10,
                                                side="bottom")

def previewBarcodes():
    
    #print(barcodes)

    previewWindow= Tk()
    
    barcodesSpread = ttk.Treeview(previewWindow)

    barcodesSpread["columns"]= (categories[1], categories[2], categories[3], categories[4])
    barcodesSpread.column("#0", width = 150)
    barcodesSpread.heading("#0", text = categories[0])
    
    for a in range(1,5):
        barcodesSpread.column(categories[a], width = 150)
        barcodesSpread.heading(categories[a], text=categories[a])

    for a in range(1,len(barcodes)):
        barcodesSpread.insert('',[a],
                              text = barcodes[a][0],
                              values=(barcodes[a][1],
                                      barcodes[a][2],
                                      barcodes[a][3],
                                      barcodes[a][4]))
        
    barcodesSpread.pack(ipadx = 5,
                        ipady = 5,
                        padx = 10,
                        pady = 10,
                        side="top")
    
    closeButton = ttk.Button(previewWindow,
                        text='Close',
                        command = previewWindow.destroy).pack(ipadx = 5,
                                                ipady = 5,
                                                padx = 10,
                                                pady = 10,
                                                side="bottom")

def clearBarcodes():
    global barcodes
    barcodes = []
    barcodes.append(categories)
    print('\n>>>')
    print('CSVBarcodes.csv | ' + (time.strftime('%c')))

############################ GUI ################################

root = Tk()
root.title('CSV Barcodes Generator')
#root.minsize(340,360)
#root.maxsize(340,360)

#################### Variables

partcodeText = tk.StringVar()
partcodeText.set('AST999999NS')
descriptionText = tk.StringVar()
descriptionText.set('13.6m Trailer')
locationText = tk.StringVar()
locationText.set('Nearside Tile')
rangeText = tk.StringVar()
rangeText.set(1)
quantityText = tk.StringVar()
quantityText.set(1)

#################### Mainframe of the GUI
mainframe = ttk.Frame(root, padding="10 10 10 10")
mainframe.pack()

#################### New barcode sub-frame
newPartcodesFrame = ttk.Labelframe(mainframe,
                                   text='New Partcode(s)',
                                   borderwidth=10,
                                   relief="solid")
newPartcodesFrame.pack()

#New barcode sub-frame - Add Button
addButton = ttk.Button(newPartcodesFrame,
                       text='+',
                       command = addNewBarcodes,
                       width = 5).grid(column=1, rowspan=5, sticky=(W),ipadx = 5, ipady = 5, padx = 5, pady = 5)

#New barcode sub-frame - Partcode subframe
partcodeFrame = ttk.Frame(newPartcodesFrame)
partcodeFrame.grid(column=2, row=1,ipadx = 2, ipady = 2, padx = 6, pady = 6)

partcodeStatic = ttk.Label(partcodeFrame, text='Partcode').pack(side='top')
partcodeInputText = ttk.Entry(partcodeFrame,width=35, textvariable = partcodeText).pack(side='bottom')

#New barcode sub-frame - Description subframe
descriptionFrame = ttk.Frame(newPartcodesFrame)
descriptionFrame.grid(column=2, row=2,ipadx = 2, ipady = 2, padx = 6, pady = 6)

descriptionStatic = ttk.Label(descriptionFrame, text='Description').pack(side='top')
descriptionInputText = ttk.Entry(descriptionFrame,width=35, textvariable = descriptionText).pack(side='bottom')

#New barcode sub-frame - Location subframe
locationFrame = ttk.Frame(newPartcodesFrame)
locationFrame.grid(column=2, row=3,ipadx = 2, ipady = 2, padx = 6, pady = 6)

locationStatic = ttk.Label(locationFrame, text='Location').pack(side='top')
locationInputText = ttk.Entry(locationFrame,width=35, textvariable = locationText).pack(side='bottom')

#Range-Quantity sub frame
rqFrame = ttk.Frame(newPartcodesFrame)
rqFrame.grid(column=2, row=4,ipadx = 10, ipady = 2, padx = 6, pady = 6)

#New barcode sub-frame - Range subframe
rangeFrame = ttk.Frame(rqFrame)
rangeFrame.pack(side='left')

rangeStatic = ttk.Label(rangeFrame, text='Range (set of)').pack(side='top')
rangeInputText = ttk.Entry(rangeFrame,width=10,textvariable = rangeText).pack(side='bottom')

#New barcode sub-frame - Quantity subframe
quantityFrame = ttk.Frame(rqFrame)
quantityFrame.pack(side='right')

quantityStatic = ttk.Label(quantityFrame, text='Quantity (of each)').pack(side='top')
quantityInputText = ttk.Entry(quantityFrame,width=10,textvariable = quantityText).pack(side='bottom')

#New barcode sub-frame - Production subframe
productionFrame = ttk.Frame(newPartcodesFrame)
productionFrame.grid(column=2, row=5,ipadx = 2, ipady = 2, padx = 6, pady = 6)

partStatic = ttk.Label(productionFrame, text='Production').pack(side='top')
var = IntVar()
var.set(1)
radio1 = ttk.Radiobutton(productionFrame, text="Print", variable=var, value=1).pack(side='left')
radio2 = ttk.Radiobutton(productionFrame, text="Cut", variable=var, value=2).pack(side='right')

#################### Load/Save files sub-frame
fileFrame = ttk.Labelframe(mainframe,
                           text='File',
                           borderwidth=10,
                           relief="solid")
fileFrame.pack()

#Load/Save  sub-frame - preview button
previewButton = ttk.Button(fileFrame,
                        text='Preview',
                        command = previewBarcodes).grid(column=1,
                                                    row=1,
                                                    ipadx = 5,
                                                    ipady = 5,
                                                    padx = 5,
                                                    pady = 5)

#Load/Save  sub-frame - clear button
clearButton = ttk.Button(fileFrame,
                        text='Clear',
                        command = clearBarcodes).grid(column=2,
                                                    row=1,
                                                    ipadx = 5,
                                                    ipady = 5,
                                                    padx = 5,
                                                    pady = 5)

#Load/Save files sub-frame - load button
loadButton = ttk.Button(fileFrame,
                        text='Load CSV',
                        command = loadCSV).grid(column=1,
                                            row=2,
                                            ipadx = 5,
                                            ipady = 5,
                                            padx = 5,
                                            pady = 5)

#Load/Save files sub-frame - save button
saveButton = ttk.Button(fileFrame,
                        text='Save CSV',
                        command = saveCSV).grid(column=2,
                                            row=2,
                                            ipadx = 5,
                                            ipady = 5,
                                            padx = 5,
                                            pady = 5)


#################### GUI main loop
root.mainloop()

#################################################################
