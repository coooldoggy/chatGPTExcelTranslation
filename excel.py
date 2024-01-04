import openai
import openpyxl


def extract_text_from_excel():
    workbook = openpyxl.load_workbook('englishA.xlsx')
    sheet = workbook.active
    text_column = sheet['A']
    text_data = [cell.value for cell in text_column if cell.value is not None]
    print(text_data)
    return text_data

def translate_text_with_gpt(text_to_translate, target_language='en'):
    prompt = f"Translate the following text to {target_language}: \n{text_to_translate}"

    openai.api_key = ''

    response = openai.Completion.create(
        engine="text-davinci-003", 
        prompt=prompt,
        max_tokens=150
    )

    translated_text = response['choices'][0]['text'].strip()
    return translated_text