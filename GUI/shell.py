from tkinter import filedialog
from tkinter import *
from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showerror
from Excel.functions import ExcelDoc


class MainFrame(Frame):
    def __init__(self):
        Frame.__init__(self)
        self.configure(bg="gray")
        self.master.minsize(150, 200)
        self.master.title("Excel Transcoder")
        self.master.rowconfigure(4, weight=1)
        self.master.columnconfigure(3, weight=1)
        self.master.resizable(width=True, height=True)
        self.master.configure(bg="gray")
        self.grid(sticky=W + E + N + S)

        self.srcText = Entry(self.master)
        self.srcText.grid(row=0, column=0, sticky=W + E + N + S, columnspan=2)
        self.srcButton = Button(self.master, text="Src File", command=self.load_file, bg="gray")
        self.srcButton.grid(row=0, column=2, sticky=W + E + N + S)
        self.destText = Entry(self.master)
        self.destText.grid(row=1, column=0, sticky=W + E + N + S, columnspan=2)
        self.destButton = Button(self.master, text="Destination", command=self.load_folder, bg = "gray")
        self.destButton.grid(row=1, column=2, sticky=W + E + N + S)

        self.exportButton = Button(self.master, text="Export", bg="gray", command=self.export)
        self.exportButton.grid(row=3, column=0, sticky=W + E + N + S, columnspan="3")


    def load_file(self):
        self.fname = askopenfilename(filetypes=[("Excel files", ".xlsx .xls .xlsm")])
        if self.fname:
            try:
                self.srcText.delete(0, END)
                self.srcText.insert(0, self.fname)
            except:  # <- naked except is a bad idea
                showerror("Open Source File", "Failed to read file\n'%s'" % self.fname)
            return
    def load_folder(self):
        self.folname = askdirectory()
        self.destText.delete(0, END)
        self.destText.insert(0, self.folname)

    def export(self):
        firstDoc = ExcelDoc(self.fname, self.folname)
        firstDoc.readFromInput()

