from tkinter import filedialog, messagebox, Entry
from ntpath import basename
import os


class Dialog:
    def __init__(self):
        self.fileNames = []
        self.pathTuple = ()
        self.folder_name = r""

    def callDialog(self):
        """Calls open file dialog, possible to choose only '.xlsx .xls .xlsm .xlsb'"""
        self.pathTuple = filedialog.askopenfilenames(filetypes=[("Excel files", ".xlsx .xls .xlsm .xlsb")])
        self.fileNames = [basename(os.path.abspath(name)) for name in self.pathTuple]
        self.folder_name = os.path.dirname(self.pathTuple[0])
    
    def callDialogOneFile(self):
        self.formatFileName = filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx .xls .xlsm .xlsb")])

    def getPaths(self):
        """Returns tuple of paths stored at class instance"""
        return self.pathTuple

    def getFolderName(self):
        return self.folder_name
    


    

