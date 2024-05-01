import openpyxl
from openpyxl.utils import get_column_letter

def sort_sinhala_words(excel_file,sheet_name,column_index):
    #Load the excel file here
    wb = openpyxl.load_workbook(excel_file)
    sheet = wb[sheet_name]

    #Get all the Sinhala words from the specified column
    sinhala_words = [cell.value for cell in sheet[get_column_letter(column_index)] if cell.value]

    #Now write the sorted value back to the excel file
    for i,word in enumerate(sorted_sinhala_words, start=1):
        sheet[get_column_letter(column_index)+str(i)] = word
