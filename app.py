import pandas as pd
import openai

openai.api_key = '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'

df = pd.read_excel('kinderschommels.xlsx')
 
columns_to_translate = ['Name', 'Category', 'Color', 'Description', 'Bullet Points']

# functie translate
def translate_text_with_openai(text):
    if pd.notnull(text):
        try:
            # Vraag OpenAI om de tekst naar Nederlands te vertalen
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "user", "content": f"Translate this text to Dutch: {text}"}
                ],
                max_tokens=350  # Limiteer de lengte van de vertaling
            )
            translated_text = response['choices'][0]['message']['content']
            return translated_text.strip()
        except Exception as e:
            print(f"Translation error: {e}")
            return text  # Geef de originele tekst terug bij een fout
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
        df[column] = df[column].apply(translate_text_with_openai)

if 'Description' in df.columns:
    df['Description'] = df['Description'].apply(improve_description)

df.to_excel('translated_and_improved_data.xlsx', index=False)

print ("Het bestand is vertaald en juist beschreven")