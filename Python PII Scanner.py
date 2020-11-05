from tkinter import filedialog
import tkinter as tk
from tkinter import ttk
from tkdatacanvas import DataCanvas
import os.path
from os import path
import re
import docx2txt
import PyPDF2
import pandas as pd
import csv
import io
import datetime
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

win = tk.Tk()
win.title("PII Scanner")

Scan_path = tk.StringVar()
Output_path = tk.StringVar()
Now = datetime.datetime.now()
CurrentDate = Now.strftime("%Y-%m-%d-%H-%M")
OutputDirectory = str(Output_path.get())
Outputlocation = OutputDirectory + CurrentDate + ".csv"

def Analytics():
    File = pd.read_csv(Outputlocation)
    PIIColumn = File['PII Type'].tolist()
    SSCount = PIIColumn.count("Social Security Number")
    CCCount = PIIColumn.count("Credit Card Number")
    PPCount = PIIColumn.count("Passport ID")
    DriversCount = PIIColumn.count("Drivers License")

    figure2 = Figure(figsize=(4, 4), dpi=70)
    pie2 = FigureCanvasTkAgg(figure2, master=win)
    pie2.get_tk_widget().grid(column=11, row=5)

    subplot2 = figure2.add_subplot(111)
    labels2 = ['Social Security Numbers', 'CreditCard Numbers', 'Passport ID Numbers', 'Drivers Licenses MD']
    pieSizes = [SSCount, CCCount, PPCount, DriversCount]
    subplot2.pie(pieSizes, autopct='%1.1f%%', shadow=True, startangle=90, radius=2)
    subplot2.legend(labels2)
    subplot2.axis('equal')
    subplot2.set_title("PII Frequency")

    PIITotal = SSCount + CCCount + PPCount + DriversCount
    MetaData = CurrentDate + ": " + str(PIITotal)
    TXTPath = OutputDirectory + "Data.txt"
    with open(TXTPath, 'a') as outputfile:
        outputfile.write(MetaData + '\n')
    with open(TXTPath, 'r') as outputfile2:
        MetaLines = [line.rstrip() for line in outputfile2]
    DateList = []
    CountList = []
    for line in MetaLines:
        MetaList = line.split(": ")
        Date = MetaList[0]
        Count = MetaList[1]
        DateList.extend(Date)
        CountList.extend(Count)


def Output(PII, Location):

    PIIlist = []
    Locationlist = []
    PIIlist.append(PII)
    Locationlist.append(Location)
    df = pd.DataFrame(data={"PII Type": PIIlist, "File Path": Locationlist})
    if os.path.isfile(Outputlocation) == True:
        df.to_csv(Outputlocation, sep=',', index=False, mode='a', header=False)
    else:
        df.to_csv(Outputlocation, sep=',', index=False)


def SSN(Data):
    if re.search(r"""\b\d{3}\-\d{2}\-\d{4}\b|\b\d{3}\s\d{2}\s\d{4}\b|Social Security Number (\b\d{9}\b)
                     |\b\d{3}\-\d{6}\b|\b\d{5}\-{4}\b""", Data):
        Type = "Social Security Number"
        Output(Type, str(filePath))


def CC(Data):
    if re.search(r"\b\d{4}\s\d{4}\s\d{4}\s\d{4}\s\b|\b\d{16}\b|\b\d{4}\-\d{4}\-\d{4}\-\d{4}\b", Data):
        Type = "Credit Card Number"
        Output(Type, str(filePath))

def PP(Data):
    if re.search(r"""Passport ID C(\w{8}\b|c\w{8}\b)|Passport Identification Number C(\w{8}\b|c\w{8}\b)|
                     Passport ID (\w{6,9}\b)|Passport Identification Number (\w{6,9}\b)""", Data):
        Type = "Passport ID"
        Output(Type, str(filePath))


def DriversMD(Data):
    if re.search(r"\b[a-zA-Z]\d{12}\b|\b[a-zA-z]\-\d{3}\-\d{3}\-\d{3}\-\d{3}\b", Data):
        Type = "Drivers License"
        Output(Type, str(filePath))


def Word(Directory):
    print(Directory)
    try:
        File = docx2txt.process(Directory)
        if SSNvar.get() == 1:
            SSN(File)
        if CCvar.get() == 1:
            CC(File)
        if PPvar.get() == 1:
            PP(File)
        if Driversvar.get() == 1:
            DriversMD(File)
    except:
        pass


def CSV(Directory):
    print(Directory)
    try:
        Data = str(pd.read_csv(Directory))
        if SSNvar.get() == 1:
            SSN(Data)
        if CCvar.get() == 1:
            CC(Data)
        if PPvar.get() == 1:
            PP(Data)
        if Driversvar.get() == 1:
            DriversMD(Data)
    except:
        pass


def Excel(Directory):
    print(Directory)
    try:
        spreadsheet = pd.ExcelFile(Directory)
        for sheet in spreadsheet.sheet_names:
            Data = str(spreadsheet.parse(sheet))
            if SSNvar.get() == 1:
                SSN(Data)
            if CCvar.get() == 1:
                CC(Data)
            if PPvar.get() == 1:
                PP(Data)
            if Driversvar.get() == 1:
                DriversMD(Data)
    except:
        pass


def Text(Directory):
    print(Directory)
    try:
        with open(Directory, mode='r') as f:
            for line in f:
                if SSNvar.get() == 1:
                    SSN(line)
                if CCvar.get() == 1:
                    CC(line)
                if PPvar.get() == 1:
                    PP(line)
                if Driversvar.get() == 1:
                    DriversMD(line)
    except:
        pass

def PDF(Directory):
    print(Directory)
    try:
        with open(Directory, mode='rb') as file:
            reader = PyPDF2.PdfFileReader(file)
            for page in reader.pages:
                PDFText = page.extractText()
                if SSNvar.get() == 1:
                    SSN(PDFText)
                if CCvar.get() == 1:
                    CC(PDFText)
                if PPvar.get() == 1:
                    PP(PDFText)
    except:
        pass



def File_sort(extensions):
    path = Scan_path.get()
    for root, directories, filenames in os.walk(path):
        for filename in filenames:
                global filePath
                filePath = os.path.join(root,filename)
                #print(filePath)
                for ext in extensions:
                    if filePath.endswith(ext):
                        if filePath.endswith(".docx") or filePath.endswith(".doc"):
                            Word(filePath)
                        elif filePath.endswith(".xlsx") or filePath.endswith(".xlx"):
                            Excel(filePath)
                        elif filePath.endswith(".pdf"):
                            PDF(filePath)
                        elif filePath.endswith(".txt"):
                            Text(filePath)
                        elif filePath.endswith(".csv"):
                            CSV(filePath)


def Scan_explorer():                                                                                                    #Opens a file explorer for the path to be scanned
    directory = filedialog.askdirectory(parent=win, initialdir="C:\\", title="File Explorer")                           #Opens the file explorer window
    Scan_path.set(directory)                                                                                            #Assigns the path to a variable
    folder1 = Scan_path.get()                                                                                           #Pulls the path from the variable
    file_path1.delete(0, tk.END)                                                                                        #Deletes previous input
    file_path1.insert(tk.END, folder1)                                                                                  #Inserts the path into the textbox
    Error_label1.configure(text="")


def Output_explorer():                                                                                                  #Opens a file explorer for the output path
    directory = filedialog.askdirectory(parent=win, initialdir="C:\\", title="File Explorer")                           #Opens the file explorer window
    Output_path.set(directory)                                                                                          #Assigns the path to a variable
    folder2 = Output_path.get()                                                                                         #Pulls the path from the variable
    file_path2.delete(0, tk.END)                                                                                        #Deletes previous input
    file_path2.insert(tk.END, folder2)                                                                                  #Inserts the path into the textbox
    Error_label2.configure(text="")


def Run():
    pd.options.display.max_columns = None
    pd.options.display.max_rows = None

    if path.exists(Scan_path.get()):
        Error_label1.configure(text="")
    else:
        Error_label1.configure(text="Path does not exist")
        Run_button.configure(state='normal')

    if path.exists(Output_path.get()):
        Error_label2.configure(text="")
    else:
        Error_label2.configure(text="Path does not exist")
        Run_button.configure(state='normal')

    file_select_count = Wordvar.get() + Excelvar.get() + Textvar.get() + PDFvar.get()
    pii_select_count = SSNvar.get() + CCvar.get() + PPvar.get()

    if file_select_count == 0:
        File_error_label.configure(text="Please make a selection")
    else:
        File_error_label.configure(text="")

    if pii_select_count == 0:
        PII_error_label.configure(text="Please make a selection")
    else:
        PII_error_label.configure(text="")

    if path.exists(Scan_path.get()) and path.exists(Output_path.get()) and file_select_count > 0 and pii_select_count > 0:

        Filelist = []
        if Wordvar.get() == 1:
            Filelist.append(".docx")
            Filelist.append(".doc")
        if Excelvar.get() == 1:
            Filelist.append(".csv")
            Filelist.append(".xlsx")
            Filelist.append(".xlx")
        if Textvar.get() == 1:
            Filelist.append(".txt")
        if PDFvar.get() == 1:
            Filelist.append(".pdf")
        File_sort(Filelist)

        with io.open(Outputlocation, "r", newline="") as csv_file:
            reader = csv.reader(csv_file)
            parsed_rows = 0
            for row in reader:
                if parsed_rows == 0:
                    dc.add_header(*row)
                else:
                    dc.add_row(*row)
                parsed_rows += 1
        dc.display()
        Run_button.configure(state='disabled', text="Finished")
    else:
        Run_button.configure(state='normal', text="Run")

    Analytics()


file_path1 = ttk.Entry(win, width=45, textvariable=Scan_path)                                                           #Entry box for the path to scan
file_path1.grid(column=1, row=1)

file_path2 = ttk.Entry(win, width=45, textvariable=Output_path)                                                         #Entry box for the output path
file_path2.grid(column=1, row=2)

Scan_label = ttk.Label(win, text="Select a path to scan: ")                                                             #Label for the scan path
Scan_label.grid(column=0, row=1)

Output_path_label = ttk.Label(win, text="Select an output path: ")                                                      #Label for the output path
Output_path_label.grid(column=0, row=2)

Browse1 = ttk.Button(win, text="Browse", command=Scan_explorer)                                                         #Browse button for the path to scan
Browse1.grid(column=2, row=1)

Browse2 = ttk.Button(win, text="Browse", command=Output_explorer)                                                       #Browse button for the output path
Browse2.grid(column=2, row=2)

Run_button = ttk.Button(win, text="Run", command=Run)
Run_button.grid(column=20, row=20)

Error_label1 = ttk.Label(win, text="")
Error_label1.grid(column=4, row=1)

Error_label2 = ttk.Label(win, text="")
Error_label2.grid(column=4, row=2)

File_types_frame = ttk.LabelFrame(win, text='File Types to be Scanned')
File_types_frame.grid(column=0, row=5, sticky=tk.N)

Wordvar = tk.IntVar()
Wordbutton = tk.Checkbutton(File_types_frame, text="Word Documents", variable=Wordvar)
Wordbutton.deselect()
Wordbutton.grid(column=0, row=0, sticky=tk.W)

Excelvar = tk.IntVar()
Excelbutton = tk.Checkbutton(File_types_frame, text="Excel Sheets", variable=Excelvar)
Excelbutton.deselect()
Excelbutton.grid(column=0, row=1, sticky=tk.W)

Textvar = tk.IntVar()
Textbutton = tk.Checkbutton(File_types_frame, text="Text Documents", variable=Textvar)
Textbutton.deselect()
Textbutton.grid(column=0, row=2, sticky=tk.W)

PDFvar = tk.IntVar()
PDFbutton = tk.Checkbutton(File_types_frame, text="PDFs", variable=PDFvar)
PDFbutton.deselect()
PDFbutton.grid(column=0, row=3, sticky=tk.W)

PII_types_frame = ttk.LabelFrame(win, text="PII Types")
PII_types_frame.grid(column=1, row=5, sticky=tk.NW)

SSNvar = tk.IntVar()
SSNbutton = tk.Checkbutton(PII_types_frame, text="Social Security Numbers", variable=SSNvar)
SSNbutton.deselect()
SSNbutton.grid(column=0, row=0, sticky=tk.W)

CCvar = tk.IntVar()
CCbutton = tk.Checkbutton(PII_types_frame, text="CreditCard Numbers", variable=CCvar)
CCbutton.deselect()
CCbutton.grid(column=0, row=1, sticky=tk.W)

PPvar = tk.IntVar()
PPbutton = tk.Checkbutton(PII_types_frame, text="Passport ID Numbers", variable=PPvar)
PPbutton.deselect()
PPbutton.grid(column=0, row=2, sticky=tk.W)

Driversvar = tk.IntVar()
Driversbutton = tk.Checkbutton(PII_types_frame, text="Drivers Licenses MD", variable=Driversvar)
Driversbutton.deselect()
Driversbutton.grid(column=0, row=3, sticky=tk.W)

File_error_label = ttk.Label(win, text="")
File_error_label.grid(column=0, row=4, sticky=tk.SW, pady=10)

PII_error_label = ttk.Label(win, text="")
PII_error_label.grid(column=1, row=4, sticky=tk.SW, pady=10)


dc = DataCanvas(win)
dc.grid(column=10, row=5)

win.mainloop()





