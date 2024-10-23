import os
from dotenv import load_dotenv
import pandas as pd
import openai

# OpenAI API key laden
load_dotenv()
openai_api_key = os.getenv('OPENAI_API_KEY')
openai.api_key = openai_api_key

df = pd.read_excel('excel/bkn-living/bknliving10.xlsx')

# Definieer de kolomnamen
column_post_title = 'post_title'
column_post_content = 'post_content'
column_meta_title = 'meta:rank_math_title'
column_meta_description = 'meta:rank_math_description'
column_focus_keyword = 'meta:rank_math_focus_keyword'
column_alg_ean = 'meta:_alg_ean'
column_global_unique_id = 'meta:_global_unique_id'
column_images = 'images'

company_name = "Bkn Living"

def improve_or_generate_focus_keyword(post_title, current_focus_keyword=None):
    try:
        # Bepaal het prompt op basis van of er al een focus keyword is
        if isinstance(current_focus_keyword, str) and len(current_focus_keyword.strip()) > 0:
            prompt = f"""
            Verbeter het volgende focus keyword en houd het binnen 60 karakters, inclusief spaties:
            '{current_focus_keyword}'.
            Zorg dat het keyword is gebaseerd op het producttype, merk, en unieke specificaties zoals maat, kleur of materiaal.
            Het product heeft de zoekwoord: '{post_title}'.
            Dit keyword zal worden gebruikt voor SEO, dus zorg dat het professioneel is en geschikt voor titels en beschrijvingen.
            """
        else:
            prompt = f"""
            Bepaal een nieuw SEO focus keyword voor het volgende product op basis van het producttype, merk en unieke specificaties zoals maat, kleur of materiaal.
            Zorg dat het keyword niet langer is dan 60 karakters, inclusief spaties.
            haal het focus keyword idee uit de kolom: '{post_title}'.
            Zorg ervoor dat het keyword professioneel is en geschikt voor titels en beschrijvingen.
            """
        
        # Vraag OpenAI om een keyword te genereren of verbeteren
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "user", "content": prompt}
            ],
            max_tokens=30  # Houd de response kort en bondig
        )
        
        # Verwerk de gegenereerde focus keyword
        focus_keyword = response['choices'][0]['message']['content'].strip()

        # Verwijder ongewenste aanhalingstekens
        focus_keyword = focus_keyword.replace('"', '').replace("'", "")

        return focus_keyword

    except Exception as e:
        print(f"Error generating focus keyword: {e}")
        return current_focus_keyword or post_title  # Als er iets misgaat, geef de huidige waarde terug
    
# Loop door de DataFrame en verbeter of genereer de focus keyword
for idx, row in df.iterrows():

    current_focus_keyword = row[column_focus_keyword] if pd.notna(row[column_focus_keyword]) else None
    # Verbeter of genereer het focus keyword
    new_focus_keyword = improve_or_generate_focus_keyword(
        post_title=row[column_post_title],
        current_focus_keyword=row[column_focus_keyword]
    )
    
    # Update de focus keyword in de bestaande kolom
    df.at[idx, column_focus_keyword] = new_focus_keyword

# Opslaan in hetzelfde bestand of een nieuw bestand als je dat wilt controleren
df.to_excel('herschreven_excel/bkn-living/updated_focus_keywords.xlsx', index=False)
