# Version 0.2
# Extract files from Zip file into new folder with the same name at the same location.

# Version 0.3
# GUI added using tkinter
# Functional extraction with select file and extract buttons


# required modules
import errno
import os
import codecs
from zipfile import ZipFile
from PIL import Image
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox


root = tk.Tk()

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
    global im1

    import_file_path = filedialog.askopenfilename()
    image1 = Image.open(import_file_path)
    im1 = image1.convert('RGB')
    image_list = []
    image_list.append(im1)
    print(image_list)


def convertImageFile():
    pass


# Button for selecting Image File(s) to convert
convertBrowseButton = tk.Button(text="Select File", command=selectImageFile, bg='green', fg='white',
                                font=('helvetica', 12, 'bold'))
canvas1.create_window(325, 140, window=convertBrowseButton)

# Button for converting selected file(s) to PDF
convertImageButton = tk.Button(text="Select File", command=convertImageFile, bg='green', fg='white',
                               font=('helvetica', 12, 'bold'))
canvas1.create_window(325, 190, window=convertImageButton)


def selectZipFile():
    global zip_file_path
    zip_file_path = filedialog.askopenfilename()


# extract function
def extractZipFile():
    # get zip file name
    zip_folder_name = zip_file_path

    # path and directory for extracted files.
    # Creates directory in the same location as ZIP file, under the same name. Removes '.zip'
    directory = zip_folder_name.replace('.zip', '')

    with ZipFile(zip_folder_name, 'r') as zip_ref:
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
extractBrowseButton = tk.Button(text="Select File", command=selectZipFile, bg='green', fg='white',
                                font=('helvetica', 12, 'bold'))
canvas1.create_window(125, 140, window=extractBrowseButton)

# Button to extract selected ZIP file
extractZipButton = tk.Button(text="Extract", command=extractZipFile, bg='green', fg='white',
                             font=('helvetica', 12, 'bold'))
canvas1.create_window(125, 190, window=extractZipButton)


if __name__ == "__main__":
    root.mainloop()
