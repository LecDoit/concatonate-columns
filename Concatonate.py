import pandas as pd                                                     # importing pandas as pd to manipulate the files
import xlsxwriter                                                       # import xlsxwriter to open file at the end
import os                                                               # import os to open the file and process
from tkinter import *                                                   # tkinter to create GUI
from tkinter import filedialog                                          # filedialog to create path selection window
# from PIL import Image, ImageTk                                        # optional - create img
import tkinter.messagebox                                               # pop up window at the end with success!

"""FUNTIONS"""
def Concatenate(event):                                                 # Defining function for later call and link it to button
        file2 = pd.read_excel(root.filename1)                           # Reading file
        df = pd.DataFrame(file2)                                        # Creating DataFrame
        lst = list(df)                                                  # Creating list of headers
        df[lst] = df[lst].astype(str)                                   # Converting Columns to strings to eliminate TimeStamps
        df['concat'] = ""                                               # Creating empty column
        for z,a in df.iterrows():                                       # Iterate all rows
                a = df.loc[z]                                           # Picking up each row and hold it as a list
                merge = "".join(a)                                      # Merge list that holding
                df.at[z, 'concat'] = merge                              # Put merged list empty column created earlier
                print(merge)                                            # Printing merged rows

        cols_order = ['concat']                                         # Prepare column to push as a first
        new_cols = cols_order + (df.columns.drop(cols_order).tolist())  # Push concat as first and add rest of the columns
        df = df[new_cols]                                               # Data Frame with new columns order
        writer = pd.ExcelWriter(root.filename2, engine='xlsxwriter')    # Store as ExcelWriter
        df.to_excel(writer, sheet_name='Sheet1', index=False)           # Convert DataFrame to Excel
        writer.save()                                                   # Export file
        print("Success")                                                # Print Success at the end
        os.startfile(root.filename2)                                    # Opening the File
        tkinter.messagebox.showinfo('Window Title',
                                    'Succes!\nConcatenated!')           # Message box pop up at the end


def FileName(event):                                                    # Defining FileName selection
        root.filename1 = filedialog.askopenfilename(initialdir = "/",
                                                    title = "Select file",
                                                    filetypes= (('xlsx files',
                                                                 "*.xlsx"),
                                                                ("all files",
                                                                 "*.*")))
        storefilename = root.filename1
        print(storefilename)

def FileSave(event):                                                    # Defining FileSave selection
        root.filename2 = filedialog.asksaveasfilename(initialdir = "/",
                                                      title = "Select file",
                                                      filetypes = (("xlsx files",
                                                                    "*.xlsx"),
                                                                   ("all files",
                                                                    "*.*")))
        storesavefilename = root.filename2
        print(storesavefilename)

"""GUI"""
root = Tk()                                                             # creating main window
root.title('Concatenation')                                             # title of main window
root.configure(background='white')                                      # setup white background

# img =Image.open("CBREjpg.png")                                        # optional - load image
# image_resized = img.resize((100,100))                                 # optional - resize
# image_resized.save('newcbre.png')                                     # optional - saving new image

CBRELogo = PhotoImage(file = 'newcbre.png')                             # optional - load image
label = Label(root,image=CBRELogo)                                      # label the image
label.grid(row=0)                                                       # put image into grid

"""Choose file button"""
FileNameButton = Button(root,text="Choose file: ", )                    
FileNameButton.bind("<Button-1>", FileName)
FileNameButton.grid(row=1)

"""Choose where to save button"""
FileSaveButton = Button(root, text="Choose where to save: ")
FileSaveButton.bind("<Button-1>", FileSave)
FileSaveButton.grid(row=2)

"""Concatenate button"""
StartButton = Button(root, text="Concatenate!")
StartButton.bind("<Button-1>", Concatenate)
StartButton.grid(row=3)

"""Quit button"""
QuitButton = Button(root, text="Quit",command=root.quit)
QuitButton.grid(row=4)

root.mainloop()                                                         # main loop to keep window open



"""
1. print box of items the rows
2. convert to exe
"""


