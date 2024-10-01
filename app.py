import pandas as pd
from googletrans import Translator
import openai

openai.api_key = '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'

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

#functie beschrijving
def improve_description(description):
    if pd.notnull(description):
        try:
            # ChatGPT wordt gevraagd om een betere beschrijving
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "user", "content": f"Verbeter deze productbeschrijving en let er op dat de specificaties ook staan beschreven: {description}"}
                ],
                max_tokens=350  # Aantal tokens in de reactie
            )
            improved_text = response['choices'][0]['message']['content']
            return improved_text.strip()
        except Exception as e:
            print(f"Error improving description: {e}")
            return description  # Geef originele tekst terug bij fout
    return description

for column in columns_to_translate:
    if column in df.columns:
        df[column] = df[column]. apply(translate_text)

if 'Description' in df.columns:
    df['Description'] = df['Description'].apply(improve_description)

df.to_excel('translated_and_improved_data.xlsx', index=False)

print ("Het bestand is vertaald en juist beschreven")