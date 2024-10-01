import os
from dotenv import load_dotenv
import pandas as pd
import openai

load_dotenv()

openai_api_key = os.getenv('OPENAI_API_KEY')

openai.api_key = openai_api_key

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
                max_tokens=700  # Limiteer de lengte van de vertaling
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
                    {"role": "user", "content": f"Verbeter deze productbeschrijving en let op dat de specificaties ook staan beschreven onder de productbeschrijving.Het moet ook in html formaat met gebruik van <p>,<h3>,<ul>,<li>. : {description}"}
                ],
                max_tokens=700  # Aantal tokens in de reactie
            )
            improved_text = response['choices'][0]['message']['content']
            return improved_text.strip()
        except Exception as e:
            print(f"Error improving description: {e}")
            return description  # Geef originele tekst terug bij fout
    return description

# functie naam verbeteren
def improve_name(name, description):
    if pd.notnull(name):
        try:
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "user", "content": f"Verbeter de productnaam '{name}' en maak het logisch en beschrijvend. Gebruik de beschrijving: '{description}' om de naam kloppend te maken."}
                ],
                max_tokens=700
            )
            improved_text = response['choices'][0]['message']['content']
            return improved_text.strip()
        except Exception as e:
            print(f"Error improving name: {e}")
            return name
    return name

# functie bullet points verbeteren
def improve_bullet_points(bullet_points):
    if pd.notnull(bullet_points):
        try:
            # Vraag OpenAI om de bullet points te verbeteren
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "user", "content": f"Verbeter de beschrijving houd het klantvriendelijk, het moet ook in html formaat met gebruik van <p>,<h3>,<ul>,<li>: {bullet_points}"}
                ],
                max_tokens=700  # Limiteer de lengte van de verbeterde bullet points
            )
            improved_text = response['choices'][0]['message']['content']
            return improved_text.strip()
        except Exception as e:
            print(f"Error improving bullet points: {e}")
            return bullet_points  # Geef de originele bullet points terug bij een fout
    return bullet_points

# Vertalen van kolommen
for column in columns_to_translate:
    if column in df.columns:
        df[column] = df[column].apply(translate_text_with_openai)

# Verbetering van de beschrijving
if 'Description' in df.columns:
    df['Description'] = df['Description'].apply(improve_description)

# Verbetering van de naam
if 'Name' in df.columns and 'Description' in df.columns:
    df['Name'] = df.apply(lambda row: improve_name(row['Name'], row['Description']), axis=1)

# Verbetering van de bullet points
if 'Bullet Points' in df.columns:
    df['Bullet Points'] = df['Bullet Points'].apply(improve_bullet_points)

# Opslaan naar een nieuw bestand
df.to_excel('improved_data.xlsx', index=False)

print ("Het bestand is vertaald en juist beschreven")