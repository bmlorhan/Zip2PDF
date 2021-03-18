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

# Version 0.4.1
# Fixed error after packaging with pyinstaller that required _cpphelpers from pikepdf

# Version 0.5
# Introduced MainApplication class. Currently all functions are nested in it, but they may be separated.
# Added custom Icon

# required modules
import errno
import os
import codecs

# ZIP extracting libraries
from zipfile import ZipFile

# Image to PDF libraries
import img2pdf
from pikepdf import _cpphelpers                             # to resolve error after packaging

# GUI libraries
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import PhotoImage


class MainApplication:
    # GUI
    def __init__(self, master):
        self.master = master
        master.title('ZIP2PDF')
        master.protocol('WM_DELETE_WINDOW', self.windowClose)

        # Canvas
        # Create window with labels above appropriate buttons
        self.canvas1 = tk.Canvas(master, width=450, height=450, bg='lightsteelblue2', relief='raised')
        self.canvas1.pack()

        # Icons
        icon1 = tk.PhotoImage(file='icons/ZIP2PDF_Logo.png')
        master.iconphoto(True, icon1)

        # Labels
        # Application Name Label
        self.appNameLabel = tk.Label(master, text='Zip2PDF', bg='lightsteelblue2')
        self.appNameLabel.config(font=('helvetica', 20))
        self.canvas1.create_window(235, 60, window=self.appNameLabel)

        # Extract Function Label
        self.extractLabel = tk.Label(master, text='Extract ZIP', bg='lightsteelblue2')
        self.extractLabel.config(font=('helvetica', 16))
        self.canvas1.create_window(125, 100, window=self.extractLabel)

        # Convert Function Label
        self.convertLabel = tk.Label(master, text='Convert to PDF', bg='lightsteelblue2')
        self.convertLabel.config(font=('helvetica', 16))
        self.canvas1.create_window(325, 100, window=self.convertLabel)

        # Buttons
        # Button for selecting Image File(s) to convert
        self.convertBrowseButton = tk.Button(text="Select Image file(s)", command=self.selectImageFile,
                                             bg='green', fg='white',font=('helvetica', 12, 'bold'))
        self.canvas1.create_window(325, 140, window=self.convertBrowseButton)

        # Button for converting selected file(s) to PDF
        self.convertImageButton = tk.Button(text="Convert Files to PDF", command=self.convertImageFile,
                                            bg='green', fg='white',font=('helvetica', 12, 'bold'))
        self.canvas1.create_window(325, 190, window=self.convertImageButton)

        # Button for selecting ZIP to extract
        self.extractBrowseButton = tk.Button(text="Select ZIP file(s)", command=self.selectZipFile,
                                             bg='green', fg='white', font=('helvetica', 12, 'bold'))
        self.canvas1.create_window(125, 140, window=self.extractBrowseButton)

        # Button to extract selected ZIP file
        self.extractZipButton = tk.Button(text="Extract ZIP file(s)", command=self.extractZipFile,
                                          bg='green', fg='white', font=('helvetica', 12, 'bold'))
        self.canvas1.create_window(125, 190, window=self.extractZipButton)

        # Close application button
        self.closeApplicationButton = tk.Button(text="Close Application", command=self.windowClose,
                                                bg='red', fg='white', font=('helvetica', 12, 'bold'))
        self.canvas1.create_window(235, 300, window=self.closeApplicationButton)

    # Functions
    # Select Image file(s) function.
    def selectImageFile(self):
        self.image_list = []
        image_file_path = filedialog.askopenfilenames()
        for images in image_file_path:
            self.image_list.append(images)

    # Convert to PDF function
    def convertImageFile(self):
        export_file_path = filedialog.asksaveasfilename(defaultextension='.pdf')
        with open(export_file_path, "wb") as f:
            f.write(img2pdf.convert(self.image_list))

    # Select ZIP file function
    def selectZipFile(self):
        zip_file_path = filedialog.askopenfilenames()
        self.zip_file_path_list = list(zip_file_path)

    # ZIP Extraction function
    def extractZipFile(self):
        # get zip file name
        zip_folder_name = self.zip_file_path_list

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

    # Window close confirmation
    def windowClose(self):
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            self.master.destroy()


def main():
    root = tk.Tk()
    app = MainApplication(root)
    root.mainloop()


if __name__ == "__main__":
    main()
