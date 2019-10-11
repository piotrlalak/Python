import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import time
import datetime
import csv
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.font_manager import FontProperties

#---------------------------------------------------------------

charsList = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','0','1','2','3','4','5','6','7','8','9','*',' ','-','$','%','.','/','+']
barsList = ['111010100010111','101110100010111','111011101000101','101011100010111','111010111000101','101110111000101','101010001110111','111010100011101','101110100011101','101011100011101','111010101000111','101110101000111','111011101010001','101011101000111','111010111010001','101110111010001','101010111000111','111010101110001','101110101110001','101011101110001','111000101010111','100011101010111','111000111010101','100010111010111','111000101110101','100011101110101','101000111011101','111010001010111','101110001010111','111011100010101','101000111010111','111010001110101','101110001110101','101000101110111','111010001011101','101110001011101','100010111011101','100011101011101','100010101110111','100010001000101','101000100010001','111000101011101','100010001010001','100010100010001']
code39 = []

in2mmUnits = 25.4
point2mmUnits = 72 / in2mmUnits

fontSizeMultiplier = 1.39665

pdfWidth = 62 / in2mmUnits # 69mm
pdfHeight = 29 / in2mmUnits # 69mm

partcodeFont = FontProperties(family='sans-serif',
                               style=None,
                               variant=None,
                               weight=None,
                               stretch=None,
                               size=(3.788 *point2mmUnits)* fontSizeMultiplier,
                               fname='C:\Windows\Fonts\Arial.ttf',
                               _init=None)

locationFont = FontProperties(family='sans-serif',
                               style=None,
                               variant=None,
                               weight=None,
                               stretch=None,
                               size=(1.768 *point2mmUnits)* fontSizeMultiplier,
                               fname='C:\Windows\Fonts\Arial.ttf',
                               _init=None)

initialDirectory = "/"
fileArray = []
categories = ['PARTCODE', 'DESCRIPTION', 'LOCATION', 'QUANTITY', 'PRODUCTION']
pdfFilename = ''
progressBarValue = 0
printItemCounter = 0
cutItemCounter = 0

#---------------------------------------------------------------

class code39Class(object):
    def __init__(self,chars,bars):
        self.chars = chars
        self.bars = bars

for a in range(0,len(charsList)):
    code39.append(code39Class(charsList[a],barsList[a]))

def combine_funcs(*funcs):
    def combined_func(*args, **kwargs):
        for f in funcs:
            f(*args, **kwargs)
    return combined_func

def fontStyle(fontSize):
    tempFontProperties = FontProperties(family='sans-serif',
                               style=None,
                               variant=None,
                               weight=None,
                               stretch=None,
                               size=(fontSize *point2mmUnits)* fontSizeMultiplier,
                               fname='C:\Windows\Fonts\Arial.ttf',
                               _init=None)
    return tempFontProperties

def rectangle(w,h,x,y,c):
    plt.fill([x,x,x+w,x+w],[y,y+h,y+h,y],c)
    
def partcode(s):
    plt.text(0.5,0.75394,s,
         horizontalalignment='center',
         verticalalignment='center',
         fontproperties = fontStyle(3.788))

def description(s):
    plt.text(0.5,0.62549,s,
         horizontalalignment='center',
         verticalalignment='center',
         fontproperties = fontStyle(1.768))
    
def location(s):
    plt.text(0.5,0.52205,s,
         horizontalalignment='center',
         verticalalignment='center',
         fontproperties = fontStyle(1.768))
    
def bars(w,h,x,y,s):
    for a in range(0,15):
        if int(s[a]) is 0:
            tempColour = 'white'
        else:
            tempColour = 'black'
        rectangle(w,h,x+(w*a),y,tempColour)

def character(w,h,x,y,s):
    rectangle(w,h,x,y,'grey')
    for a in range(0,len(charsList)):
        if s is charsList[a]:
            bars(w,h,x,y,barsList[a])
    
def barcode(w,h,x,y,s):
    s ='*'+s+'*'
    barW = w / ( (len(s)*15) + (len(s)-1) )
    charW =  barW *15
    for a in range(0,len(s)):
        character(barW,h,x,y,s[a])
        x=x+charW
        if a < (len(s)-1):
            rectangle(barW,h,x,y,'white')
            x=x+barW

def label(partcodeStr,descriptionStr,locationStr):
    rectangle(1,1,0,0,'white')
    partcode(partcodeStr)
    description(descriptionStr)
    location(locationStr)
    barcode(0.91936,0.25862,0.04032,0.15517,partcodeStr)#size and position is fixed

def full_frame():
    import matplotlib as mpl
    mpl.rcParams['savefig.pad_inches'] = 0
    fig = plt.figure(figsize=(pdfWidth,pdfHeight))
    ax = plt.axes([0,0,1,1], frameon=False)
    ax.get_xaxis().set_visible(False)
    ax.get_yaxis().set_visible(False)
    plt.autoscale(tight=True)

def getFilenamePath(filename):
    tempDirectory = filename.split('/')
    currentDirectory = ''
    for a in range(0,(len(tempDirectory)-1)):
        if tempDirectory[a] == '':
            currentDirectory += '/'
        else:
            currentDirectory += tempDirectory[a] + '/'
    return currentDirectory

def loadCSV():
    global initialDirectory,pdfFilename,fileArray,printItemCounter,cutItemCounter
    fileArray = []
    barcodeListLength.set('0')
    printItemCounter = 0
    cutItemCounter = 0
    pb["maximum"] = 0
    filename = tk.filedialog.askopenfilename(filetypes=(('CSV Files','.csv'),('All files', '*')))
    if filename != '':
        csvData = csv.reader(open(filename))
        initialDirectory = getFilenamePath(filename)
    else:
        initialDirectory = '/'
    for row in csvData:
        if row == categories:
            pass
        else:
            fileArray.append(row)
            if row[4] == 'Print' or row[4] == 'PRINT':
                printItemCounter  += 1
            else:
                cutItemCounter += 1
                
                    
    if len(fileArray) > 0:
        pdfFilename = (filename.split('/')[len(filename.split('/'))-1]).split('.')[0]
        csvFileName.set((filename.split('/')[len(filename.split('/'))-1]).split('.')[0])
        barcodeListLength.set((str(len(fileArray)))
                              + ' (' + str(printItemCounter) + ' print, '
                              + str(cutItemCounter) + ' cut)')
        firstBarcode.set(fileArray[0][0])
        pb["maximum"] = len(fileArray)

def saveFile():
    global pdfFilename,fileArray,progressBarValue
    filename = tk.filedialog.asksaveasfilename(defaultextension = '.pdf',
                                               initialdir = initialDirectory,
                                               initialfile = pdfFilename,
                                               title = "Save file",
                                               filetypes = (("PDF file","*.pdf"),("All files","*.*")))
    if not filename:
        pass
    else:
        if len(fileArray) > 0:
            exportLabels(filename)
        else:
            pass
    #clear-out for next file
    fileArray = []
    csvFileName.set('Load CSV File')
    barcodeListLength.set('0')
    firstBarcode.set('N/A')
    progressBarValue = 0
    pb["value"] = progressBarValue
    pb["maximum"] = len(fileArray)

def singleLabel(array):
    global progressBarValue
    full_frame()
    label(array[a][0],
        array[a][1],
        array[a][2],)
    pdf.savefig()
    plt.close()
    progressBarValue+=1
    pb["value"] = progressBarValue

def exportLabels(filename):
    global progressBarValue
    with PdfPages(filename) as pdf:
        #Main label generation
        for a in range(0,len(fileArray)):
            if var.get() is 1:
                if fileArray[a][4] == 'Print' or fileArray[a][4] == 'PRINT':
                    progressBarValue+=1
                    pass
                else:
                    #singleLabel(fileArray)
                    full_frame()
                    label(fileArray[a][0],
                        fileArray[a][1],
                        fileArray[a][2],)
                    pdf.savefig()
                    plt.close()
                    progressBarValue+=1
                    pb["value"] = progressBarValue
            else:
                #singleLabel(fileArray)
                full_frame()
                label(fileArray[a][0],
                    fileArray[a][1],
                    fileArray[a][2],)
                pdf.savefig()
                plt.close()
                progressBarValue+=1
                pb["value"] = progressBarValue
                
        #Metadata
        d = pdf.infodict()
        d['Title'] = (filename.split('/')[len(filename.split('/'))-1]).split('.')[0]
        d['Author'] = u'Piotr Lalak'
        d['Subject'] = '62x29mm Barcode Labels'
        d['Keywords'] = '62x29mm Barcode Labels ' + (filename.split('/')[len(filename.split('/'))-1]).split('.')[0]
        d['CreationDate'] = datetime.datetime.today()
        d['ModDate'] = datetime.datetime.today()

#---------------------------------------------------------------

root = Tk()
root.title('62x29mm Labels')
#root.minsize(340,360)
#root.maxsize(340,360)   

csvFileName = StringVar()
csvFileName.set('Load CSV File')
barcodeListLength = StringVar()
barcodeListLength.set('0')
firstBarcode = StringVar()
firstBarcode.set('N/A')

mainframe = ttk.Frame(root, padding="10 10 10 10")
mainframe.pack()

fileSpecsFrame = ttk.Labelframe(mainframe,text='Info', borderwidth=10, relief="solid")
fileSpecsFrame.pack()

static1 = ttk.Label(fileSpecsFrame, text='CSV File: ').grid(column=1, row=1, sticky=(W))
static2 = ttk.Label(fileSpecsFrame, text='No. of Barcodes: ').grid(column=1, row=2, sticky=(W))
static3 = ttk.Label(fileSpecsFrame, text='First Barcode: ').grid(column=1, row=3, sticky=(W))
loadedFileLabel = ttk.Label(fileSpecsFrame, textvariable=csvFileName).grid(column=2, row=1, sticky=(E))
barcodeListLabel = ttk.Label(fileSpecsFrame, textvariable=barcodeListLength).grid(column=2, row=2, sticky=(E))
firstBarcodeLabel = ttk.Label(fileSpecsFrame, textvariable=firstBarcode).grid(column=2, row=3, sticky=(E))

pb = ttk.Progressbar(mainframe, orient=HORIZONTAL,
                     length=200, mode='determinate',
                     maximum = len(fileArray),
                     value = progressBarValue)
pb.pack(ipadx = 5, ipady = 5,padx = 5, pady = 5)

var = IntVar()
var.set(1)
skipPrintBox = ttk.Checkbutton(mainframe,
                               text="Skip print items",
                               variable=var).pack(ipadx = 5,
                                             ipady = 5,
                                             padx = 5,
                                             pady = 5)

fileFrame = ttk.Labelframe(mainframe,
                           text='File',
                           borderwidth=10,
                           relief="solid")
fileFrame.pack()
loadButton = ttk.Button(fileFrame,
                    text='Load CSV',
                    command = loadCSV).pack(ipadx = 5,
                                            ipady = 5,
                                            padx = 5,
                                            pady = 5,
                                            side="left")
saveButton = ttk.Button(fileFrame,
                    text='Save PDF',
                    command = saveFile).pack(ipadx = 5,
                                            ipady = 5,
                                            padx = 5,
                                            pady = 5,
                                            side="right")

firstBarcodeLabel = ttk.Label(mainframe, text='Version 4 - 5th June 2019').pack()
root.mainloop()

