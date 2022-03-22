# @Time: 01.03.2022 14:26
# @Author: Jia Wang
# @File: numpy.py
# @Software: PyCharm


import os
import win32com.client as win32


# Folder Path
path = 'C:/Users/Data/Test'
# only select file extension with .xlsx
ext = ('.xlsx')

# Iterate through all the files and put them in a list
def get_filelist(dir):
    Filelist = []
    for home, dirs,files in os.walk(path):
        for filename in files:
            if filename.endswith(ext):
                Filelist.append(os.path.join(home,filename))

    return Filelist

# Check if file is being used
def is_used(file):
    if os.path.exists(file):
        try:
            os.rename(file, file)
            return False
        except:
            return True
    raise NameError

# open-refresh-save file
def open_close_excel(file):

    try:
        Xlsx = win32.DispatchEx('Excel.Application')
        Xlsx.DisplayAlerts = False
        Xlsx.Visible = False
        book = Xlsx.Workbooks.Open(file)
        book.RefreshAll()
        print(file, ": is being Refreshed.....")
        Xlsx.CalculateUntilAsyncQueriesDone()
        book.Save()
        book.Close(SaveChanges=True)
        Xlsx.Quit()

        book = None
        Xlsx = None
        del book
        del Xlsx
        print("Refresh Done!")
    except Exception as e:
        print(e)

    finally:
        book = None
        Xlsx = None

if __name__ == "__main__":

    Filelist = get_filelist(dir)

    #Filter out all files in Archive,alt, and DataSource folders

    list_string = ['_Archiv', 'Archiv', 'alt', 'DataSource']
    filtered_filelist = list(filter(lambda text: all([word not in text for word in list_string]),Filelist))
    for file in filtered_filelist:
        if is_used(file) == True:
            print(file, "is being used, you can not refresh it!")
        else:
            open_close_excel(file)
