from tkinter import *
from tkinter import filedialog
from tkinter.filedialog import askopenfilename,asksaveasfile
from PyPDF2 import PdfFileReader

# =================open file method======================
def openFile():
    file = askopenfilename(defaultextension=".xlsx",
                           filetypes=[("Excel Worksheet", "*.xlsx")])
    if file == "":
        file = None
    else:
        fileEntry.delete(0, END)
        fileEntry.config(fg="blue")
        fileEntry.insert(0, file)


def convert():
    import student




#=================== Front End Design
root = Tk()
root.geometry("600x350")
root.config(bg="light blue")
root.title(" Result analysis [ designed by mrc ] ")
root.resizable(0, 0)
try:
    root.wm_iconbitmap("pdf2.ico")
except:
    print('icon file is not available')
    pass
file = ""
defaultText = "\n\n\n\n\t\t DATTA MEGHE COLLEGE OF ENGINEERING .\n \t\t     RESULT ANALYSIS SYSTEM."

# ==============App Name==============================================================>>
appName = Label(root, text=" RESULT ANALYSIS ", font=('arial Black ', 20),
                bg="light blue", fg='maroon')
appName.place(x=150, y=5)
# Select pdf file
labelFile = Label(root, text="Select file", font=('arial', 12, 'bold'))
labelFile.place(x=30, y=70)
fileEntry = Entry(root, font=('calibri', 12), width=20)
fileEntry.pack(ipadx=200, pady=50, padx=150)
# ===========button to access openFile method=================================
openFileButton = Button(root, text=" Choose an excel file ", font=('arial', 12, 'bold'), width=30,
                        bg="sky blue", fg='green', command=openFile)
openFileButton.place(x=150, y=80)
# ===========button to access convert method=================================
convert2Text = Button(root, text="Read", font=('arial', 8, 'bold'),
                      bg="RED", fg='WHITE', width=20, command=convert)
convert2Text.place(x=250, y=120)
# ======================= Text Box to read pdf file and modify ===================>>
readPdf = Text(root, font=('calibri', 12), fg='light green', bg='black', width=60, height=30, bd=10)
readPdf.pack(padx=20, ipadx=20, pady=20, ipady=20)
readPdf.insert(INSERT, defaultText)



# ===================halt window=============================>>
if __name__ == "__main__":
    root.mainloop()
    convert()