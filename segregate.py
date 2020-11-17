import os
import time
import zipfile

# Global Variables
path_to_watch = 'C:/Users/kungf/Desktop/Statement/Files/Downloads'
folder_destination = 'C:/Users/kungf/Desktop/Statement'


def ds_extract(name, s_type, date, password):
    file_path = folder_destination + '/Zip/' + s_type + '/' + date + '.zip'

    with zipfile.ZipFile(file_path) as file:
        file.extractall(pwd=bytes(password, 'utf-8'))
        try:
            os.rename(folder_destination + '/Code/' + name + '.xls',
                      folder_destination + '/Files/' + s_type + '/' + date + '.xls')
        except FileExistsError:
            print('File already exists : ' + folder_destination +
                  '/Files/' + s_type + '/' + date + '.xls')


def extract_zip(name, s_type, date, password):
    """
    Extract Zip files

    :s_type: Type of Statement
    :date: Date of file
    :password: Password to open the file
    """
    if s_type == 'DS' or s_type == 'US':
        ds_extract(name, s_type, date, password)


def movefile(file: str):
    """
    Function to move files to right location

    :file: File name
    """
    name, ext = os.path.splitext(file)
    # if ext == '.xls':
    #     try:
    #         copyfile(path_to_watch + '/' + file, folder_destination + '/' + file)
    #     except FileExistsError:
    #         print('File already exists : ' + folder_destination + '/' + file)
    if ext == '.zip':
        if name.startswith('037011000390216'):
            password, s_type, date = name.split('_')

            try:
                if s_type == 'STATEMENT':
                    return
                os.rename(path_to_watch + '/' + file, folder_destination +
                          '/Zip/' + s_type + '/' + date + ext)
                extract_zip(name, s_type, date, password)
            except FileExistsError:
                print('File already exists : ' + folder_destination +
                      '/Zip/' + s_type + '/' + date + ext)


def main():
    """
    Main runner Function
    """
    before = dict([(f, None) for f in os.listdir(path_to_watch)])

    for i in before:
        movefile(i)

    print('Initial Load Complete ...')

    while 1:
        time.sleep(15)
        after = dict([(f, None) for f in os.listdir(path_to_watch)])
        added = [f for f in after if not f in before]
        if added:
            print("Added: ", ", ".join(added))
        for i in added:
            movefile(i)
        before = after

        print('Cycle Complete ... ')


if __name__ == '__main__':
    main()
