import openpyxl
import openai

from excel import extract_text_from_excel, translate_text_with_gpt

input_file = 'englishA.xlsx'
output_file = 'empty.xlsx'

def translate_excel_file():
    text_to_translate = extract_text_from_excel()
    translated_text = [translate_text_with_gpt(text) for text in text_to_translate]

    workbook = openpyxl.Workbook()
    sheet = workbook.active

    for i, translation in enumerate(translated_text):
        sheet.cell(row=i+1, column=1, value=translation)

    workbook.save(output_file)


translate_excel_file()
