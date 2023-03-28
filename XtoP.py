import win32com.client
from pywintypes import com_error
import os
from PyPDF2 import PdfReader, PdfMerger, PdfWriter

path_to_excels_folders = input('Enter the path to the folder containing the Excel files: ').replace('/','\\')

path_to_save_PDFs = input('Enter the path to the folder for the output PDF files: ').replace('/','\\')

if path_to_save_PDFs == '':
    path_to_save_PDFs = path_to_excels_folders

# Get the list of all the files in the directory
files = os.listdir(path_to_excels_folders)


count_notConverted = 0
notConverted = []

count_converted = 0
converted = []

input("Press ENTER to start conversion")

#Iterate through all the Exxcel files in the folder
for i in files:
    #Getting the file name without the extension (.xlsx/.xls)
    if os.path.isfile(os.path.join(path_to_excels_folders, i)):
        file_name_without_extension, file_extension = os.path.splitext(i)
        
    try:
        # Path to original excel file
        PATH_TO_EXCEL = fr'{path_to_excels_folders}\\{i}'
       
        # PDF path when saving
        PATH_TO_PDF = fr'{path_to_save_PDFs}\\{file_name_without_extension}.pdf'
        
        # Opening Excel
        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = False
        workbook = excel.Workbooks.Open(PATH_TO_EXCEL)
       
       # Conversion and Saving
        print(f'Starting conversion to PDF of {file_name_without_extension}')
        
        # Excel sheet you want to work in
        workbook.Worksheets[1]

        workbook.ActiveSheet.ExportAsFixedFormat(0, PATH_TO_PDF)

    # Error Handling
    except com_error as error:
        print(error)
        count_notConverted +=1
        notConverted.append(i)
        pass

    else:
        print(f'Succesfully converted {i} to PDF')
        count_converted += 1
        converted.append(i)
    
    finally:
        # Close Workbook and Excel
        workbook.Close()
        excel.Quit()

print(f'There were {count_notConverted} documents that were not converted: {notConverted}')
print(f'There were {count_converted} documents that were succesfully converted: {converted}')



