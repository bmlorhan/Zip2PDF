""" ZIP2PDF
    The primary functions are to allow the user to extract selected ZIP files
    and then convert image files to a single PDF."""


# required modules
import errno
import os
import codecs

# ZIP extracting libraries
from zipfile import ZipFile

# GUI libraries
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
# Image to PDF libraries
import img2pdf

# PDF merging libraries
from PyPDF2 import PdfFileMerger

# This import is specifically to address an error post packaging
from pikepdf import _cpphelpers


# Main Application class
class MainApplication:
    """ Builds out the GUI """

    # pylint: disable=too-many-instance-attributes
    # 12 is reasonable in this case
    # pylint: disable=line-too-long

    def __init__(self, master):
        self.master = master
        master.title('ZIP2PDF')
        master.protocol('WM_DELETE_WINDOW', self.window_close)

        # Canvas
        # Create window with labels above appropriate buttons
        self.canvas1 = tk.Canvas(master, width=450, height=475, bg='lightsteelblue2', relief='raised')
        self.canvas1.pack()

        # Labels
        # Application Name Label
        self.app_name_label = tk.Label(master, text='Zip2PDF', bg='lightsteelblue2')
        self.app_name_label.config(font=('helvetica', 20))
        self.canvas1.create_window(235, 40, window=self.app_name_label)

        # Extract Function Label
        self.extract_label = tk.Label(master, text='Extract ZIP', bg='lightsteelblue2')
        self.extract_label.config(font=('helvetica', 16))
        self.canvas1.create_window(125, 100, window=self.extract_label)

        # Convert Function Label
        self.convert_label = tk.Label(master, text='Convert to PDF', bg='lightsteelblue2')
        self.convert_label.config(font=('helvetica', 16))
        self.canvas1.create_window(325, 100, window=self.convert_label)

        # Buttons
        # Button for selecting Image File(s) to convert
        self.convert_browse_button = tk.Button(text="Select Image file(s)", command=self.select_image_file,
                                               bg='green', fg='white', font=('helvetica', 12, 'bold'))
        self.canvas1.create_window(325, 140, window=self.convert_browse_button)

        # Button for converting selected file(s) to PDF
        self.convert_image_button = tk.Button(text="Convert Files to PDF", command=self.convert_image_file,
                                              bg='green', fg='white', font=('helvetica', 12, 'bold'))
        self.canvas1.create_window(325, 190, window=self.convert_image_button)

        # Button for converting selected file(s) to PDF and combining them with existing PDF
        self.convert_and_combine_button = tk.Button(text="Combine PDF files", command=self.combine_pdf_files,
                                                    bg='green', fg='white', font=('helvetica', 12, 'bold'))
        self.canvas1.create_window(325, 240, window=self.convert_and_combine_button)

        # Button for selecting ZIP to extract
        self.extract_browse_button = tk.Button(text="Select ZIP file(s)", command=self.select_zip_file,
                                               bg='green', fg='white', font=('helvetica', 12, 'bold'))
        self.canvas1.create_window(125, 140, window=self.extract_browse_button)

        # Button to extract selected ZIP file
        self.extract_zip_button = tk.Button(text="Extract ZIP file(s)", command=self.extract_zip_file,
                                            bg='green', fg='white', font=('helvetica', 12, 'bold'))
        self.canvas1.create_window(125, 190, window=self.extract_zip_button)

        # Close application button
        self.close_application_button = tk.Button(text="Close Application", command=self.window_close,
                                                  bg='red', fg='white', font=('helvetica', 12, 'bold'))
        self.canvas1.create_window(235, 400, window=self.close_application_button)

    # Functions
    # Select Image file(s) function.
    def select_image_file(self):
        """ User selects the image file(s) they wish to convert to PDF"""
        self.image_list = []
        image_file_path = filedialog.askopenfilenames()
        for images in image_file_path:
            self.image_list.append(images)

    # Convert to PDF function
    def convert_image_file(self):
        """ Converts the selected images and merges them into a single PDF at the desired location and name"""
        export_file_path = filedialog.asksaveasfilename(defaultextension='.pdf')
        with open(export_file_path, "wb") as pdf_file:
            pdf_file.write(img2pdf.convert(self.image_list))

    @staticmethod                                      # static until further notice
    def combine_pdf_files():
        """ Combines PDF files to create a single file"""
        merger = PdfFileMerger()
        selected_pdfs_list = []
        selected_pdfs = filedialog.askopenfilenames()

        # adds selected PDFs to list for later use
        for pdfs in selected_pdfs:
            selected_pdfs_list.append(pdfs)

        # select file name and save location of final PDF output
        final_pdf_file_path = filedialog.asksaveasfilename(defaultextension='.pdf')
        with open(final_pdf_file_path, 'wb') as final_pdf:
            for pdf_files in selected_pdfs_list:
                merger.append(pdf_files)
            merger.write(final_pdf)

    # Select ZIP file function
    def select_zip_file(self):
        """ User selects the ZIP file(s) they wish to extract via Windows Explorer"""
        zip_file_path = filedialog.askopenfilenames()
        self.zip_file_path_list = list(zip_file_path)

    # ZIP Extraction function
    def extract_zip_file(self):
        """ Extracts the user selected file(s). Saves them to the same location as the ZIP file(s) with the same name"""
        # get zip file name
        zip_folder_name = self.zip_file_path_list

        for file_path in zip_folder_name:
            # path and directory for extracted files.
            # Creates directory in the same location as ZIP file, under the same name. Removes '.zip'
            directory = file_path.replace('.zip', '')
            with ZipFile(file_path, 'r') as zip_ref:
                for files_in_zip in zip_ref.infolist():
                    bad_filename = files_in_zip.filename
                    # tries to encode using cp437 first
                    try:
                        decoded_files = bad_filename.encode('cp437').decode('sjis')
                    # if cp437 doesn't work, tries cp932. Mainly for Katakana
                    except UnicodeEncodeError:
                        print('uf did not decode to "sjis", attempting to decode to sjis/ ')   # shift_jisx0213
                        decoded_files = bad_filename.encode('cp932').decode('sjis')

                    # Saves unzipped file to the same location as original ZIP file.
                    filename = os.path.join(directory, decoded_files)
                    if not os.path.exists(os.path.dirname(filename)):
                        try:
                            os.makedirs(os.path.dirname(filename))
                        except OSError as exc:
                            if exc.errno != errno.EEXIST:
                                raise

                    if not filename.endswith('/'):
                        with open(filename, 'wb') as dest:
                            dest.write(zip_ref.read(files_in_zip))

    # Window close confirmation
    def window_close(self):
        """ Basic windows close confirmation message."""
        if messagebox.askokcancel("Quit", "Do you want to close the application?"):
            self.master.destroy()


# Application
def main():
    """ application loop """
    root = tk.Tk()
    root.resizable(False, False)  # prevent window resizing
    MainApplication(root)
    root.mainloop()


if __name__ == "__main__":
    main()
