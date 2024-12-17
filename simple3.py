import os
import re
from dotenv import load_dotenv
import pandas as pd
from openai import OpenAI
import time
import json

load_dotenv()

openai_api_key = os.getenv('OPENAI_API_KEY2')

# Initialiseer de OpenAI client
client = OpenAI(api_key=openai_api_key)

# Laad de DataFrame
df = pd.read_excel('herschreven_excel/simpledeal/homcomproducten_1.1.xlsx', dtype={'meta:_alg_ean': str})

column_post_name = 'post_name'
column_focus_keyword = 'meta:_yoast_wpseo_focuskw'
column_meta_title = 'meta:_yoast_wpseo_title'
column_meta_description = 'meta:_yoast_wpseo_metadesc'
column_ean = 'meta:_alg_ean'
column_gtin = 'meta:wpseo_global_identifier_values'

company_name = "Simpledeal"

df[column_meta_title] = df[column_meta_title].astype(str)
df[column_meta_description] = df[column_meta_description].astype(str)
df[column_post_name] = df[column_post_name].astype(str)
df[column_ean] = df[column_ean].astype(str)
df[column_gtin] = df[column_gtin].astype(str)

# Functie voor het herschrijven van de URL (post_name)
def rewrite_post_name(focus_keyword, post_name, max_retries=3):
    prompt = f"""
    Herschrijf de URL '{post_name}' zodat deze begint met het focus keyword '{focus_keyword}', 
    post name bevat maximaal 4-5 woorden en niet meer dan 70 karakters inclusief spaties. Gebruik alleen letters, cijfers, en koppeltekens.
    Geen speciale tekens of slashes (/). Ook geen afmetingen of cijfers. 
    Vertaal de url van het Duits naar correct Nederlands. 
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

        except Exception as e:
            print(f"Error rewriting post name (attempt {attempt + 1}): {e}")
            time.sleep(1)  

    # Als alle pogingen falen, geef het originele post_name terug
    return post_name

# Functie voor het genereren van de meta title
def generate_meta_title(focus_keyword):
    try:
        prompt = (
            f"Schrijf een SEO-geoptimaliseerde meta title, beginnend met het focus keyword '{focus_keyword}', en gevolgd door een powerword. "
            "De titel moet bestaan uit 5-6 woorden, en mag niet langer zijn dan 60 karakters inclusief spaties. "
            "Zorg ervoor dat de titel niet wordt afgebroken en aantrekkelijk is voor de lezer."
        )

        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=30
        )

        meta_title = response.choices[0].message.content.strip().replace('"', '').replace("'", "")
        
        if len(meta_title) < 47:
            full_title = f"{meta_title} | {company_name}"
        else:
            full_title = meta_title

        return full_title
    except Exception as e:
        print(f"Error generating meta title: {e}")
        return f"{focus_keyword} | {company_name}"


def generate_meta_description(post_title, focus_keyword):
    try:
        prompt = f"""
        Schrijf een SEO-geoptimaliseerde meta description beginnend met het focus keyword '{focus_keyword}'. 
        De meta description moet aantrekkelijk zijn, ongeveer 20 woorden bevatten. precies 2 zinnen bevatten en tussen 120 en 150 tekens inclusief spaties lang zijn.
        De meta description moet eindeigen met een punt. Dus geen uitroepteken of vraagteken. 
        Gebruik de '{post_title}' als idee.  
        De meta description moet aantrekkelijk zijn en exact 2 zinnen bevatten:
        - De eerste zin beschrijft het product, beginnend met het focus keyword '{focus_keyword}'.
        - De tweede zin eindigt met een duidelijke call-to-action.
        """

        # Vraag een meta description van de AI
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=200
        )

        # Verwerk de output van de AI
        meta_description = response.choices[0].message.content.strip()

        # Splits zinnen op punten, vraagtekens of uitroeptekens gevolgd door spatie
        sentences = re.split(r'(?<=[.!?])\s+', meta_description)

        # Valideer de zinnen en beperk tot maximaal 2
        if len(sentences) > 2:
            sentences = sentences[:2]  # Houd alleen de eerste twee zinnen

        # Combineer de eerste twee zinnen en zorg dat het eindigt met een punt
        meta_description = ' '.join(sentences).strip()
        if not meta_description.endswith('.'):
            meta_description += '.'

        # Corrigeer ongewenste combinaties zoals !. of ?. door deze te vervangen met een enkelvoudig symbool
        meta_description = re.sub(r'[!?]+\.', '.', meta_description)

         # Voeg een derde zin toe als de lengte onder 120 tekens ligt
        if len(meta_description) <= 120:
            meta_description += " Bestel nu bij Simpledeal."

        return meta_description

    except Exception as e:
        print(f"Error generating meta description: {e}")
        return f"Product beschrijving van {post_title} met focus op {focus_keyword}."

def format_gtin_value(ean_value):
    print(f"Processing EAN value: {ean_value}")  # Dit helpt bij het debuggen

    ean_value = str(ean_value).strip()  # Converteer de waarde naar string en verwijder extra spaties

    # Verwijder de '.0' als het een float-achtige waarde is
    if ean_value.endswith(".0"):
        ean_value = ean_value[:-2]  # Verwijder de laatste twee tekens

    if not ean_value:
        return json.dumps({
            "gtin8": "",
            "gtin12": "",
            "gtin13": "",
            "gtin14": "",
            "isbn": "",
            "mpn": ""
        })
    
    # Zet de waarde altijd in het gtin13 veld
    gtin13 = ean_value

    # Format de resulterende waarde als JSON-structuur
    result = {
        "gtin8": "",
        "gtin12": "",
        "gtin13": gtin13,
        "gtin14": "",
        "isbn": "", 
        "mpn": ""    
    }
    
    time.sleep(1)  # Voeg een korte vertraging toe, indien nodig
    # Zet het resultaat om naar een JSON-string
    return json.dumps(result)

# Verwerking van de DataFrame
for index, row in df.iterrows():
    print(f"Processing row {index + 1} of {len(df)}")
    
    post_name = row[column_post_name]
    focus_keyword = row[column_focus_keyword]
    
    new_post_name = rewrite_post_name(focus_keyword, post_name)
    df.at[index, column_post_name] = new_post_name

    new_meta_title = generate_meta_title(focus_keyword)
    df.at[index, column_meta_title] = new_meta_title

    new_meta_description = generate_meta_description(row['post_title'], focus_keyword)
    df.at[index, column_meta_description] = new_meta_description
    
    # Verwerken van EAN en het geformatteerde GTIN-resultaat
    ean_value = row[column_ean]  
    formatted_value = format_gtin_value(ean_value)  

    # Voeg de geformatteerde waarde toe aan de meta:wpseo_global_identifier_values kolom
    df.at[index, column_gtin] = formatted_value

# Opslaan
output_file = 'herschreven_excel/simpledeal/homcomproducten_final_1.xlsx'
df.to_excel(output_file, index=False)

print("Verwerking voltooid! Resultaten zijn opgeslagen in:", output_file)