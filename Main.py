from tkinter import filedialog
import tkinter as tk
from tkinter import ttk
from tkdatacanvas import DataCanvas
from tkinter import scrolledtext
from os import path
import csv
import io
import datetime
import glob
from piiPatternMatching import *
from fileTypes import *


class PiiScanner:

    # Initializes variables and the GUI
    def __init__(self, master):
        self.scanPath = tk.StringVar()
        self.outputPath = tk.StringVar()
        self.Now = datetime.datetime.now()
        self.currentDate = self.Now.strftime("%Y-%m-%d-%H-%M")
        self.outputDirectory = str(self.outputPath.get())
        self.outputLocation = self.outputDirectory + "/" + self.currentDate + ".csv"

        self.master = master
        self.master.title("PII Scanner")

        self.wordVar = tk.IntVar()
        self.excelVar = tk.IntVar()
        self.textVar = tk.IntVar()
        self.pdfVar = tk.IntVar()

        self.ssnVar = tk.IntVar()
        self.ccVar = tk.IntVar()
        self.driversVar = tk.IntVar()

        self.console = scrolledtext.ScrolledText(self.master, state="normal", wrap=tk.WORD, width=120, height=10)
        self.runButton = ttk.Button(self.master, text="Run", command=self.run)
        self.filePath1 = ttk.Entry(self.master, width=45, textvariable=self.scanPath)
        self.filePath2 = ttk.Entry(self.master, width=45, textvariable=self.outputPath)
        self.scanLabel = ttk.Label(self.master, text="Select a path to scan: ")
        self.outputPathLabel = ttk.Label(self.master, text="Select an output path: ")
        self.Browse1 = ttk.Button(self.master, text="Browse", command=self.scan_explorer)
        self.Browse2 = ttk.Button(self.master, text="Browse", command=self.output_explorer)
        self.errorLabel1 = ttk.Label(self.master, text="")
        self.errorLabel2 = ttk.Label(self.master, text="")
        self.fileTypesFrame = ttk.LabelFrame(self.master, text='File Types to be Scanned')
        self.piiErrorLabel = ttk.Label(self.master, text="")
        self.fileErrorLabel = ttk.Label(self.master, text="")
        self.piiTypesFrame = ttk.LabelFrame(self.master, text="PII Types")
        self.dc = DataCanvas(self.master)

        self.wordButton = tk.Checkbutton(self.fileTypesFrame, text="Word Documents", variable=self.wordVar)
        self.excelButton = tk.Checkbutton(self.fileTypesFrame, text="Excel Sheets", variable=self.excelVar)
        self.textButton = tk.Checkbutton(self.fileTypesFrame, text="Text Documents", variable=self.textVar)
        self.pdfButton = tk.Checkbutton(self.fileTypesFrame, text="PDFs", variable=self.pdfVar)
        self.ssnButton = tk.Checkbutton(self.piiTypesFrame, text="Social Security Numbers", variable=self.ssnVar)
        self.ccButton = tk.Checkbutton(self.piiTypesFrame, text="CreditCard Numbers", variable=self.ccVar)
        self.driversButton = tk.Checkbutton(self.piiTypesFrame, text="Drivers Licenses MD", variable=self.driversVar)

        self.wordButton.select()
        self.excelButton.select()
        self.textButton.select()
        self.pdfButton.select()
        self.ssnButton.select()
        self.ccButton.select()
        self.driversButton.select()

        self.fileTypesFrame.grid(column=0, row=5, padx=20, sticky=tk.W)
        self.piiTypesFrame.grid(column=1, row=5, sticky=tk.NW)

        self.console.grid(column=0, row=7, columnspan=20, pady=10, padx=20)
        self.filePath1.grid(column=1, row=1, sticky=tk.W)
        self.filePath2.grid(column=1, row=2, sticky=tk.W)
        self.scanLabel.grid(column=0, row=1, padx=20, sticky=tk.W)
        self.outputPathLabel.grid(column=0, row=2, padx=20, sticky=tk.W)
        self.Browse1.grid(column=2, row=1)
        self.Browse2.grid(column=2, row=2)
        self.runButton.grid(column=20, row=20)
        self.errorLabel1.grid(column=4, row=1)
        self.errorLabel2.grid(column=4, row=2)
        self.fileErrorLabel.grid(column=0, row=4, sticky=tk.SW, pady=10)
        self.piiErrorLabel.grid(column=1, row=4, sticky=tk.SW, pady=10)
        self.dc.grid(column=0, row=6, columnspan=2, pady=10, padx=20, sticky=tk.W)
        self.wordButton.grid(column=0, row=0, sticky=tk.W)
        self.excelButton.grid(column=0, row=1, sticky=tk.W)
        self.textButton.grid(column=0, row=2, sticky=tk.W)
        self.pdfButton.grid(column=0, row=3, sticky=tk.W)
        self.ssnButton.grid(column=0, row=0, sticky=tk.W)
        self.ccButton.grid(column=0, row=1, sticky=tk.W)
        self.driversButton.grid(column=0, row=3, sticky=tk.W)

    # Executes when the "browse" button is clicked by the user next to the scan box. Opens a file explorer window
    def scan_explorer(self):
        directory = filedialog.askdirectory(parent=win, initialdir="C:\\",
                                            title="File Explorer")
        self.scanPath.set(directory)
        folder1 = self.scanPath.get()
        self.filePath1.delete(0, tk.END)
        self.filePath1.insert(tk.END, folder1)
        self.errorLabel1.configure(text="")

    # Executes when the "browse" button is clicked by the user next to the output box. Opens a file explorer window
    def output_explorer(self):
        directory = filedialog.askdirectory(parent=win, initialdir="C:\\",
                                            title="File Explorer")
        self.outputPath.set(directory)
        folder2 = self.outputPath.get()
        self.filePath2.delete(0, tk.END)
        self.filePath2.insert(tk.END, folder2)
        self.errorLabel2.configure(text="")
        self.outputDirectory = str(self.outputPath.get())
        self.outputLocation = self.outputDirectory + "/" + self.currentDate + ".csv"

    # Processes Social Security Numbers from the data and prepares the data for output
    def ssn_process(self, data, filename):
        ssn_pii = ssn(str(data))
        for pii in ssn_pii:
            pii_type = "Social Security Number"
            self.output(pii_type, filename, pii)

    # Processes Credit Cards from the data and prepares the data for output
    def cc_process(self, data, filename):
        cc_pii = cc(str(data))
        for pii in cc_pii:
            pii_type = "Credit Card Number"
            self.output(pii_type, filename, pii)

    # Processes Drivers License IDs from the data and prepares the data for output
    def drivers_process(self, data, filename):
        drivers_pii = drivers_md(str(data))
        for pii in drivers_pii:
            pii_type = "Drivers License MD"
            self.output(pii_type, filename, pii)

    # Populates the Output table with the data exported to the csv file
    def output_table(self):
        with io.open(self.outputLocation, "r", newline="") as csv_file:
            reader = csv.reader(csv_file)
            parsed_rows = 0
            for row in reader:
                if parsed_rows == 0:
                    self.dc.add_header(*row)
                else:
                    self.dc.add_row(*row)
                parsed_rows += 1
            self.dc.display()

    # Checks if all needed parameters are present and displays errors if they are not
    def input_sanitation(self):
        file_select_count = self.wordVar.get() + self.excelVar.get() + self.textVar.get() + self.pdfVar.get()
        pii_select_count = self.ssnVar.get() + self.ccVar.get()

        if path.exists(self.scanPath.get()):
            self.errorLabel1.configure(text="")
        else:
            self.errorLabel1.configure(text="Path does not exist")
            self.runButton.configure(state='normal')
            return False

        if path.exists(self.outputPath.get()):
            self.errorLabel2.configure(text="")
        else:
            self.errorLabel2.configure(text="Path does not exist")
            self.runButton.configure(state='normal')
            return False

        if file_select_count == 0:
            self.fileErrorLabel.configure(text="Please make a selection")
            return False
        else:
            self.fileErrorLabel.configure(text="")

        if pii_select_count == 0:
            self.fileErrorLabel.configure(text="Please make a selection")
            return False
        else:
            self.fileErrorLabel.configure(text="")
        return True

    # Prepares the data to be exported via CSV
    def output(self, pii_type, location, pii):
        pii_list = []
        pii_type_list = []
        location_list = []
        pii_type_list.append(pii_type)
        location_list.append(location)
        pii_list.append(pii)
        df = pd.DataFrame(data={"PII Type": pii_type_list, "File Path": location_list, "PII": pii_list})
        if path.isfile(self.outputLocation):
            df.to_csv(self.outputLocation, sep=',', index=False, mode='a', header=False)
        else:
            df.to_csv(self.outputLocation, sep=',', index=False)

    # Updates the live console. (this is what makes it live)
    def console_update(self, filename):
        self.console.insert(tk.END, filename + "\n")
        self.console.yview(tk.END)
        self.master.update()

    # Executes when the user clicks the "run" button. Acts as main
    def run(self):
        data = ""
        scan_path = self.scanPath.get()
        pd.options.display.max_columns = None
        pd.options.display.max_rows = None

        if self.input_sanitation():
            for filename in glob.iglob(scan_path + '/**', recursive=True):
                is_extension = False
                if filename != self.outputLocation:
                    if self.wordVar.get() == 1 and filename.endswith(".docx") or filename.endswith(".doc"):
                        self.console_update(filename)
                        data = word_sort(filename)
                        is_extension = True

                    elif self.excelVar.get() == 1 and filename.endswith(".xlsx") or filename.endswith(".xlx"):
                        self.console_update(filename)
                        data = excel_sort(filename)
                        is_extension = True

                    elif self.pdfVar.get() == 1 and filename.endswith(".pdf"):
                        self.console_update(filename)
                        data = pdf_sort(filename)
                        is_extension = True

                    elif self.textVar.get() == 1 and filename.endswith(".txt"):
                        self.console_update(filename)
                        data = text_sort(filename)
                        is_extension = True

                    elif self.excelVar.get() == 1 and filename.endswith(".csv"):
                        self.console_update(filename)
                        data = csv_sort(filename)
                        is_extension = True

                if is_extension:

                    if self.ssnVar.get() == 1:
                        self.ssn_process(data, filename)

                    if self.ccVar.get() == 1:
                        self.cc_process(data, filename)

                    if self.driversVar.get() == 1:
                        self.drivers_process(data, filename)

            self.output_table()
            self.runButton.configure(state='disabled', text="Finished")
        else:
            self.runButton.configure(state='normal', text="Run")


win = tk.Tk()
gui = PiiScanner(win)
win.mainloop()

