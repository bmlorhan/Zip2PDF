# Version 0.2
# Extract files from Zip file into new folder with the same name at the same location.

# Version 0.3
# GUI added using tkinter.
# Functional extraction with select file and extract buttons.

# Version 0.3.1
# ZIP extraction can now select and extract multiple files.
# Added Exit Application button.

# Version 0.4
# Primary features are complete.
# JPG files can now be selected and converted to a single PDF file at desired location with desired name.
# Added window close confirmation messagebox.


# required modules
import errno
import os
import codecs
from zipfile import ZipFile
import img2pdf
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox


root = tk.Tk()

# Create window with labels above appropriate buttons
canvas1 = tk.Canvas(root, width=450, height=450, bg='lightsteelblue2', relief='raised')
canvas1.pack()

# Application Name Label
appNameLabel = tk.Label(root, text='Zip2PDF', bg='lightsteelblue2')
appNameLabel.config(font=('helvetica', 20))
canvas1.create_window(235, 60, window=appNameLabel)

# Extract Function Label
extractLabel = tk.Label(root, text='Extract ZIP', bg='lightsteelblue2')
extractLabel.config(font=('helvetica', 16))
canvas1.create_window(125, 100, window=extractLabel)

# Convert Function Label
convertLabel = tk.Label(root, text='Convert to PDF', bg='lightsteelblue2')
convertLabel.config(font=('helvetica', 16))
canvas1.create_window(325, 100, window=convertLabel)


# Select File function.
def selectImageFile():
    global image_list
    image_list = []
    image_file_path = filedialog.askopenfilenames()
    for images in image_file_path:
        image_list.append(images)


def convertImageFile():
    export_file_path = filedialog.asksaveasfilename(defaultextension='.pdf')
    with open(export_file_path, "wb") as f:
        f.write(img2pdf.convert(image_list))


# Button for selecting Image File(s) to convert
convertBrowseButton = tk.Button(text="Select Image file(s)", command=selectImageFile, bg='green', fg='white',
                                font=('helvetica', 12, 'bold'))
canvas1.create_window(325, 140, window=convertBrowseButton)

# Button for converting selected file(s) to PDF
convertImageButton = tk.Button(text="Convert Files to PDF", command=convertImageFile, bg='green', fg='white',
                               font=('helvetica', 12, 'bold'))
canvas1.create_window(325, 190, window=convertImageButton)


def selectZipFile():
    global zip_file_path_list
    zip_file_path_list = []
    zip_file_path = filedialog.askopenfilenames()
    zip_file_path_list = list(zip_file_path)


# extract function
def extractZipFile():
    # get zip file name
    zip_folder_name = zip_file_path_list

    for file_path in zip_folder_name:
        # path and directory for extracted files.
        # Creates directory in the same location as ZIP file, under the same name. Removes '.zip'
        directory = file_path.replace('.zip', '')
        with ZipFile(file_path, 'r') as zip_ref:
            for f in zip_ref.infolist():
                bad_filename = f.filename
                if bytes != str:
                    bad_filename = bytes(bad_filename, 'cp437')

                # decode to sjis
                try:
                    uf = codecs.decode(bad_filename, 'sjis')
                except UnicodeDecodeError:
                    print('uf did not decode to "sjis", attempting to decode to shift_jisx0213. ')
                    uf = codecs.decode(bad_filename, 'shift_jisx0213')

                filename = os.path.join(directory, uf)
                if not os.path.exists(os.path.dirname(filename)):
                    try:
                        os.makedirs(os.path.dirname(filename))
                    except OSError as exc:
                        if exc.errno != errno.EEXIST:
                            raise

                if not filename.endswith('/'):
                    with open(filename, 'wb') as dest:
                        dest.write(zip_ref.read(f))


# Button for selecting ZIP to extract
extractBrowseButton = tk.Button(text="Select ZIP file(s)", command=selectZipFile, bg='green', fg='white',
                                font=('helvetica', 12, 'bold'))
canvas1.create_window(125, 140, window=extractBrowseButton)

# Button to extract selected ZIP file
extractZipButton = tk.Button(text="Extract ZIP file(s)", command=extractZipFile, bg='green', fg='white',
                             font=('helvetica', 12, 'bold'))
canvas1.create_window(125, 190, window=extractZipButton)


# Window close confirmation
def windowClose():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        root.destroy()


# Close application button
closeApplicationButton = tk.Button(text="Close Application", command=windowClose, bg='red', fg='white',
                                   font=('helvetica', 12, 'bold'))
canvas1.create_window(235, 300, window=closeApplicationButton)


if __name__ == "__main__":
    root.protocol("WM_DELETE_WINDOW", windowClose)
    root.mainloop()
