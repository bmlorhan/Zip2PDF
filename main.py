# Version 0.2
# Extract files from Zip file into new folder with the same name.

# required modules
import errno
from zipfile import ZipFile
import os
import codecs


# extract function
def extract():
    # get zip file name
    zip_folder_name = input("enter file path: ")

    # path and directory for extracted files. Creates directory in the same location as ZIP file, under the same name.
    # removes '.zip'
    directory = zip_folder_name.replace('.zip', '')

    # create if non-existent
    # if not os.path.exists(directory):
    #     os.makedirs(directory)

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

            # print files to screen
            print(repr(uf))

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

        print("Extraction complete.")


if __name__ == "__main__":
    extract()
