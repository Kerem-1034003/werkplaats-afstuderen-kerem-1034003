import pandas as pd
from googletrans import Translator

translator = Translator()

df = pd.read_excel('kinderschommels.xlsx')
 
columns_to_translate = ['Name', 'Category', 'Color', 'Description', 'Bullet Points']

# functie translate
def translate_text(text):
    if pd.notnull(text):
        try: 
            translated_text = translator.translate(text, src='de', dest='nl').text
            return translated_text
        except Exception as e: 
            print(f"vertaling error:{e}")
        return text
    return text

for column in columns_to_translate:
    if column in df.columns:
        df[column] = df[column]. apply(translate_text)

df.to_excel('translated_and_improved_data.xlsx', index=False)

print ("Het bestand is vertaald en juist beschreven")