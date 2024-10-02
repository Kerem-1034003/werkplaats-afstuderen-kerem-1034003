import os
from dotenv import load_dotenv
import pandas as pd
import openai

load_dotenv()

openai_api_key = os.getenv('OPENAI_API_KEY')

openai.api_key = openai_api_key

df = pd.read_excel('kinderschommels.xlsx')
columns_low_temp = ['Material','Category', 'Color']
columns_high_temp = ['Name','Description','Bullet Points']

abbreviations = ['PE', 'PU', 'PVC'] 

# functie voor vertalen met lage temperature (voor nauwkeurigheid)
def translate_text_low_temp(text):
    if pd.notnull(text):
        if text in abbreviations:
            return text
        try:
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "user", "content": f"Translate this text to Dutch: {text}"}
                ],
                temperature=0,  # Lage temperature voor consistentie
                max_tokens=700
            )
            translated_text = response['choices'][0]['message']['content']
            return translated_text.strip()
        except Exception as e:
            print(f"Translation error (low temp): {e}")
            return text
    return text

# functie voor vertalen met hogere temperature (voor vloeiende output)
def translate_text_high_temp(text):
    if pd.notnull(text):
        try:
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "user", "content": f"Translate this text to Dutch: {text}"}
                ],
                temperature=0.7,  # Hogere temperature voor vloeiendere vertaling
                max_tokens=700
            )
            translated_text = response['choices'][0]['message']['content']
            return translated_text.strip()
        except Exception as e:
            print(f"Translation error (high temp): {e}")
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
                    {"role": "user", "content": f"Verbeter de productnaam '{name}' en maak het logisch. Gebruik de beschrijving: '{description}' om de naam kloppend te maken."}
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

# Stap 1: Vertaling van de kolommen met lage temperature
for column in columns_low_temp:
    if column in df.columns:
        df[column] = df[column].apply(translate_text_low_temp)

# Stap 2: Vertaling van de kolommen met hogere temperature (Description, Bullet Points)
for column in columns_high_temp:
    if column in df.columns:
        df[column] = df[column].apply(translate_text_high_temp)

# Stap 3: Verbetering van de beschrijving
if 'Description' in df.columns:
    df['Description'] = df['Description'].apply(improve_description)

# Stap 4: Verbetering van de naam met behulp van description
if 'Name' in df.columns and 'Description' in df.columns:
    df['Name'] = df.apply(lambda row: improve_name(row['Name'], row['Description']), axis=1)

# Stap 5: Verbetering van de bullet points
if 'Bullet Points' in df.columns:
    df['Bullet Points'] = df['Bullet Points'].apply(improve_bullet_points)

# Opslaan naar een nieuw bestand
df.to_excel('improved_data.xlsx', index=False)

print ("Het bestand is vertaald en juist beschreven")