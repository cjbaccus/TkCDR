#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
TkCDR
Author: Carl Baccus

This program is designed to provide a GUI for my PyCDR
program to clean Cisco Communications Mangaer (CUCM) Call Detail Records (CDR's) and write them out to an excel spreadsheet

Future enhancements: write out pivot charts in new sheets to get stats for calls.

Notes for local installations:
# Mac
The following requirements must be installed for this program to work
1. Python version 2.7X
2. PIP
    a. curl https://bootstrap.pypa.io/ez_setup.py -o - | sudo python
    b. sudo easy_install pip
3. From PIP install the following libraries
    a. pip install xlsxwritere
    b. pip install csv
    c. pip install Tkinter
    d. pip install re
    e. pip install datetime

"""


from Tkinter import Frame, Tk, BOTH, Text, Menu, END
import tkFileDialog 
import sys
import csv
import time
import re
import xlsxwriter
from datetime import datetime
 

class Example(Frame):
  
    def __init__(self, parent):
        Frame.__init__(self, parent)   
         
        self.parent = parent        
        self.initUI()
        
        
    def initUI(self):
      
        self.parent.title("CDR Report Generator")
        self.pack(fill=BOTH, expand=1)
        
        menubar = Menu(self.parent)
        self.parent.config(menu=menubar)
        
        fileMenu = Menu(menubar)
        fileMenu.add_command(label="Raw CDR", command=self.onOpen)
        menubar.add_cascade(label="Choose CDR to Clean.", menu=fileMenu) 
        fileMenu.add_command(label="Exit", command=self.onExit)       
        
        self.txt = Text(self)
        self.txt.pack(fill=BOTH, expand=1)

    def readFile(self, filename):

        f = open(filename, "r")
        text = f.read()
        return text

    def date_and_time(self, time_value):
        return time.strftime("%m/%d/%y %H:%M:%S", time.localtime(float(time_value)))

    def convert_duration(self, secs):
        secs = int(secs)
        m, s = divmod(secs, 60)
        h, m = divmod(m, 60)
        return "%d:%02d:%02d" % (h, m, s)

    def onExit(self):
        self.quit()

    def onOpen(self):

        workbook = xlsxwriter.Workbook('CleanedCDR.xlsx')
        worksheet = workbook.add_worksheet()
        
        # set formatting for date in excel
        format1 = workbook.add_format({'num_format': 'MM/DD/YY hh:mm:ss'})
        ftypes = [('CDR files', '*.txt'), ('All files', '*')]
        dlg = tkFileDialog.Open(self, filetypes = ftypes)
        fl = dlg.show()
        
        with open(fl, "r") as infile:
            reader = csv.reader(infile)
            next(reader, None)  # skip the headers
            xrow = 0
            xcol = 0
            self.txt.insert(END, "Reading CDR file...\n")
            worksheet.write(xrow, xcol, "Date-time")
            worksheet.write(xrow, xcol + 1, "Duration")
            worksheet.write(xrow, xcol + 2, "Calling Number")
            worksheet.write(xrow, xcol + 3, "Called Number")
            worksheet.write(xrow, xcol + 4, "Final Called Number")
            worksheet.write(xrow, xcol + 5, "finalCalled-UserID")
            xrow = 1
            for row in reader:
                if row[47] == "0":
                    pass
                elif re.match("5000", row[29]):
                    def convert_duration(secs):
                        secs = int(secs)
                        m, s = divmod(secs, 60)
                        h, m = divmod(m, 60)
                        return "%d:%02d:%02d" % (h, m, s)
                    as_datetime = datetime.strptime(time.strftime("%m/%d/%y %H:%M:%S", time.localtime(float((row[47])))), '%m/%d/%y %H:%M:%S')
                    worksheet.write(xrow, xcol, as_datetime, format1)
                    worksheet.write(xrow, xcol + 1, convert_duration(row[55]))
                    worksheet.write(xrow, xcol + 2, row[8])
                    worksheet.write(xrow, xcol + 3, row[29])
                    worksheet.write(xrow, xcol + 4, row[30])
                    worksheet.write(xrow, xcol + 5, row[31])
                    xrow += 1
        self.txt.insert(END, "All Done...\n\nImported and excel sheet created!")

def main():
  
    root = Tk()
    ex = Example(root)
    root.geometry("300x250+300+300")
    root.mainloop()  


if __name__ == '__main__':
    main()  