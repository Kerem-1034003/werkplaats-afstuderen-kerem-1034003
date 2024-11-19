import os
from dotenv import load_dotenv
import pandas as pd
from openai import OpenAI
import time
import re

# Laad de OpenAI API key
load_dotenv()
openai_api_key = os.getenv('OPENAI_API_KEY')

# Initialiseer de OpenAI client
client = OpenAI(api_key=openai_api_key)

df = pd.read_excel ('herschreven_excel/bkn-living/updated_bknliving1.4.xlsx')

# Definieer de kolomnamen
column_post_name = 'post_name'
column_focus_keyword = 'meta:rank_math_focus_keyword'
column_alg_ean = 'meta:_alg_ean'
column_global_unique_id = 'meta:_global_unique_id'
column_images = 'images'

# Functie voor het herschrijven van de URL (post_name)
def rewrite_post_name(focus_keyword, post_name, max_retries=3):
    prompt = f"""
    Herschrijf de URL '{post_name}' zodat deze begint met het focus keyword '{focus_keyword}', 
    bevat 5-6 woorden en niet meer dan 70 karakters inclusief spaties. Gebruik alleen letters, cijfers, en koppeltekens.
    Geen speciale tekens of slashes (/). 
    """

    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=100  # Beperk tot 50 tokens voor een compacte URL
            )
            # Verwijder ongewenste tekens en controleer lengte
            new_post_name = response.choices[0].message.content.strip().replace('"', '').replace("'", "")
            
            # Gebruik regex om alle tekens behalve alfanumerieke karakters en koppeltekens te verwijderen
            new_post_name = re.sub(r"[^a-zA-Z0-9\- ]", "", new_post_name)
            
            # Vervang spaties door koppeltekens en maak de tekst geschikt voor URL
            new_post_name = new_post_name.replace(" ", "-")

            # Controleer of de URL binnen de limiet van 70 karakters valt
            if len(new_post_name) <= 70:
                return new_post_name
            else:
                # Als de naam nog steeds te lang is, beperk tot de eerste 5 woorden
                words = new_post_name.split("-")
                new_post_name = "-".join(words[:5])

                # Controleer opnieuw de lengte
                if len(new_post_name) <= 70:
                    return new_post_name
                else:
                    print(f"Warning: Generated URL still too long after adjustment: '{new_post_name}'")

        except Exception as e:
            print(f"Error rewriting post name (attempt {attempt + 1}): {e}")
            time.sleep(2)  # Wacht 2 seconden voor een nieuwe poging

    # Als alle pogingen falen, geef het originele post_name terug
    return post_name

# Functie om EAN naar global_unique_id te kopiëren
def copy_ean_to_global_unique_id(df, ean_column='meta:_alg_ean', global_unique_id_column='meta:_global_unique_id'):
    for idx, row in df.iterrows():
        ean_value = row[ean_column]
        if pd.notna(ean_value):
            df.at[idx, global_unique_id_column] = ean_value
    return df

# Functie voor het toevoegen van alt-tekst aan afbeeldingen
def add_alt_text_to_images(df, images_column='images', focus_keyword_column='meta:rank_math_focus_keyword', max_retries=3):
    for idx, row in df.iterrows():
        if pd.notna(row[images_column]):
            images = row[images_column].split(' | ')
            if images:
                prompt = f"""
                Genereer een korte en beschrijvende alt-tekst van ongeveer 75 karakters voor een productafbeelding.
                Gebruik het volgende focus keyword: '{row[focus_keyword_column]}'. 
                Houd de beschrijving relevant en helder, en in correct Nederlands.
                """
                
                for attempt in range(max_retries):
                    try:
                        response = client.chat.completions.create(
                            model="gpt-3.5-turbo",
                            messages=[{"role": "user", "content": prompt}],
                            max_tokens=60,
                            temperature=0.7
                        )
                        
                        alt_text = response.choices[0].message.content.strip()
                        parts = images[0].split(' ! ')
                        new_parts = []
                        for part in parts:
                            if part.startswith('alt :'):
                                new_parts.append(f"alt : {alt_text}")
                            else:
                                new_parts.append(part)
                        images[0] = ' ! '.join(new_parts)
                        break
                    
                    except Exception as e:
                        print(f"Error generating alt text for row {idx} on attempt {attempt + 1}: {e}")
                        time.sleep(2)
                
                df.at[idx, images_column] = ' | '.join(images)

    return df

# Verwerk elke rij in de DataFrame voor herschrijven van post_name
for index, row in df.iterrows():
    print(f"Processing row {index + 1} of {len(df)}")
    
    # URL herschrijven
    focus_keyword = row[column_focus_keyword]
    post_name = row[column_post_name]
    new_post_name = rewrite_post_name(focus_keyword, post_name)
    df.at[index, column_post_name] = new_post_name

    # Pauze na elke API-aanroep om de rate limits te respecteren
    time.sleep(1)

    # Sla tussentijdse resultaten op na elke 50 rijen om verlies van gegevens te voorkomen
    if (index + 1) % 50 == 0:
        temp_output_file = f'herschreven_excel/bkn-living/temp_output_{index + 1}.xlsx'
        df.to_excel(temp_output_file, index=False)
        print(f"Tussentijdse resultaten opgeslagen in: {temp_output_file}")

# Alt-tekst toevoegen aan afbeeldingen en EAN kopiëren naar Global Unique ID
df = add_alt_text_to_images(df)
df = copy_ean_to_global_unique_id(df)

# Sla de volledige resultaten op in een nieuw Excel-bestand
output_file = 'herschreven_excel/bkn-living/updated_bknliving4.xlsx'
df.to_excel(output_file, index=False)

print("Verwerking voltooid! Resultaten zijn opgeslagen in:", output_file)
