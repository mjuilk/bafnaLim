from tkinter import *
from tkinter import filedialog
from tkinter import ttk
import subprocess

root = Tk()

def change_theme():
    # NOTE: The theme's real name is azure-<mode>
    if root.tk.call("ttk::style", "theme", "use") == "azure-dark":
        # Set light theme
        root.tk.call("set_theme", "light")
    else:
        # Set dark theme
        root.tk.call("set_theme", "dark")

def browseFile():
    filePath = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    fileEntry.delete(0, END)
    fileEntry.insert(0, filePath)

def convertFile():
    filePath = fileEntry.get()
    subprocess.call(["python", "animalTrackingExtractPDF.py", filePath])

# Pack a big frame so, it behaves like the window background
frame1 = ttk.Frame(root)
frame1.pack(expand = True)

frame2 = ttk.Frame(root)
frame2.pack(expand = True)

# Set the initial theme
root.tk.call("source", "azure.tcl")
root.tk.call("set_theme", "dark")

root.title("PDF to CSV Converter")
root.minsize(640, 240)
root.maxsize(1920, 1080)

fileLabel = Label(frame1, text="Select PDF file to convert:", justify = 'center')
fileLabel.grid(row = 0, column = 0)

fileEntry = Entry(frame1, width=50)
fileEntry.grid(row = 1, column = 0)

browseButton = ttk.Button(frame2, text = "Browse", style = 'Accent.TButton', command=browseFile)
browseButton.grid(row = 2, column = 0, padx = 10)

convertButton = ttk.Button(frame2, text="Convert", style = 'Accent.TButton', command=convertFile)
convertButton.grid(row = 2, column = 1, padx = 10)

root.mainloop()