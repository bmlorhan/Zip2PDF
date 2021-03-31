""" ZIP2PDF
    The primary functions are to allow the user to extract selected ZIP files
    and then convert image files to a single PDF."""

# required modules
import errno
import os

# ZIP extracting libraries
from zipfile import ZipFile
from unrar import rarfile
from py7zr import SevenZipFile

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
        # Button for selecting individual image file(s) to convert
        self.convert_browse_button = tk.Button(text="Select Image file(s)", command=self.select_image_file,
                                               bg='green', fg='white', font=('helvetica', 12, 'bold'))
        self.canvas1.create_window(325, 140, window=self.convert_browse_button)

        # Button for selecting image folder(s) to convert all images inside.
        self.convert_browse_image_folder_button = tk.Button(text='Select Image Folder(s)',
                                                            command=self.select_image_folder,
                                                            bg='green', fg='white', font=('helvetica', 12, 'bold'))
        self.canvas1.create_window(325, 290, window=self.convert_browse_image_folder_button)

        # Button for converting selected file(s) to PDF
        self.convert_image_button = tk.Button(text="Convert Files to PDF", command=self.convert_image_file,
                                              bg='green', fg='white', font=('helvetica', 12, 'bold'))
        self.canvas1.create_window(325, 190, window=self.convert_image_button)

        # Button for converting selected file(s) to PDF and combining them with existing PDF
        self.convert_and_combine_button = tk.Button(text="Combine PDF files", command=self.combine_pdf_files,
                                                    bg='green', fg='white', font=('helvetica', 12, 'bold'))
        self.canvas1.create_window(325, 240, window=self.convert_and_combine_button)

        # Button for selecting ZIP to extract
        self.extract_browse_button = tk.Button(text="Select Archive file(s)", command=self.select_archive_file,
                                               bg='green', fg='white', font=('helvetica', 12, 'bold'))
        self.canvas1.create_window(125, 140, window=self.extract_browse_button)

        # Button to extract selected ZIP file
        self.extract_zip_button = tk.Button(text="Extract Archive file(s)", command=self.extract_archive_file,
                                            bg='green', fg='white', font=('helvetica', 12, 'bold'))
        self.canvas1.create_window(125, 190, window=self.extract_zip_button)

        # Close application button
        self.close_application_button = tk.Button(text="Close Application", command=self.window_close,
                                                  bg='red', fg='white', font=('helvetica', 12, 'bold'))
        self.canvas1.create_window(235, 400, window=self.close_application_button)

        # Global Variables used in various functions
        # Used in select_image_file and convert_image_file
        self.image_list = []

        # Used in select_archive_file and extract_archive_file
        self.zip_file_path_list = []

    # Functions
    # Select Image file(s) function.
    def select_image_file(self):
        """ User selects the image file(s) they wish to convert to PDF"""
        image_file_path = filedialog.askopenfilenames()
        for images in image_file_path:
            self.image_list.append(images)

    # Convert to PDF function
    def convert_image_file(self):
        """ Converts the selected images and merges them into a single PDF at the desired location and name"""
        export_file_path = filedialog.asksaveasfilename(defaultextension='.pdf')
        with open(export_file_path, "wb") as pdf_file:
            pdf_file.write(img2pdf.convert(self.image_list))

    # Select image folder(s) function and convert all image files inside to PDF
    @staticmethod
    def select_image_folder():
        """ Converts ALL image files within the selected image folder to a single PDF"""
        image_folder_path = filedialog.askdirectory()
        image_folder_save_path = filedialog.asksaveasfilename(defaultextension='pdf')
        with open(image_folder_save_path, 'wb') as save_folder:
            image_list = []
            for file_name in os.listdir(image_folder_path):
                if not file_name.endswith('.jpg'):
                    continue
                path = os.path.join(image_folder_path, file_name)
                if os.path.isdir(path):
                    continue
                image_list.append(path)
            save_folder.write(img2pdf.convert(image_list))

    # Combine PDF files function
    @staticmethod
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
    def select_archive_file(self):
        """ User selects the ZIP file(s) they wish to extract via Windows Explorer"""
        zip_file_path = filedialog.askopenfilenames()
        self.zip_file_path_list = list(zip_file_path)

    # ZIP Extraction function
    def extract_archive_file(self):
        """ Extracts the user selected file(s).
        Saves them to the same location as the original archive with the same name"""
        # Dictionary to store file extensions and their appropriate extraction function
        archive_dict = {
            '.zip': ZipFile,
            '.rar': rarfile.RarFile,
            '.7z': SevenZipFile,
        }

        # get zip file name
        zip_folder_name = self.zip_file_path_list
        for file_path in zip_folder_name:

            # get file name and extension.
            # Extension is used to use the right filetype library function for extraction.
            file_name, file_extension = os.path.splitext(file_path)
            with archive_dict[file_extension](file_path, 'r') as archive_ref:

                # This try loop is because SevenZipFile's attribute is named differently than ZipFile and RarFile
                try:
                    for files_in_archive in archive_ref.infolist():
                        bad_filename = files_in_archive.filename

                        self.save_extractions(bad_filename, file_name, archive_ref, files_in_archive)

                # An AttributeError is raised, so the exception uses SevenZipFile's attribute.
                # This may be a problem if more libraries are supported and also have different attribute names
                except AttributeError:
                    for files_in_archive in archive_ref.getnames():
                        bad_filename = archive_ref.filename
                        archive_ref.reset()
                        self.save_extractions(bad_filename, file_name, archive_ref, files_in_archive)

    # Save files
    def save_extractions(self, bad_filename, file_name, archive_ref, files_in_archive):

        # Runs encode_decode function for Japanese Kanji, Hiragana, and Katakana
        decoded_files = self.encode_decode_function(bad_filename)
        # Saves unzipped file to the same location as original ZIP file.
        final_file_name = os.path.join(file_name, decoded_files)
        if not os.path.exists(os.path.dirname(final_file_name)):
            try:
                os.makedirs(os.path.dirname(final_file_name))
            except OSError as exc:
                if exc.errno != errno.EEXIST:
                    raise

        if not final_file_name.endswith('/'):
            with open(final_file_name, 'wb') as dest:

                # Similar try loop as in extract_zip_file
                try:
                    dest.write(archive_ref.read(files_in_archive))

                # SevenZipFile extraction works differently from ZipFile and RarFile
                # and uses different attribute names"
                except TypeError:
                    archive_ref.extractall(file_name)

    # Encode / Decode function
    @staticmethod
    def encode_decode_function(bad_filename):
        """ Function tries to encode to 'cp437' and decode to 'sjis'.
        If it cannot be encoded to 'cp437' it will attempt to use 'cp932'. This is primarily for Katakana"""
        # tries to encode using cp437 first
        try:
            decoded_files = bad_filename.encode('cp437').decode('sjis')
        # if cp437 doesn't work, tries cp932. Mainly for Katakana
        except UnicodeEncodeError:  # shift_jisx0213
            decoded_files = bad_filename.encode('cp932').decode('sjis')
        return decoded_files

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
