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
df = pd.read_excel('../herschreven_excel/simpledeal/simpledeal/script1_part1.xlsx', dtype={'meta:_alg_ean': str})

column_post_name = 'post_name'
column_focus_keyword = 'meta:_yoast_wpseo_focuskw'
column_meta_title = 'meta:_yoast_wpseo_title'
column_meta_description = 'meta:_yoast_wpseo_metadesc'
column_ean = 'meta:_alg_ean'
column_gtin = 'meta:wpseo_global_identifier_values'
column_post_title = 'post_title'

company_name = "Simpledeal"

df[column_meta_title] = df[column_meta_title].astype(str)
df[column_meta_description] = df[column_meta_description].astype(str)
df[column_post_name] = df[column_post_name].astype(str)
df[column_ean] = df[column_ean].astype(str)
df[column_gtin] = df[column_gtin].astype(str)
df[column_post_title] = df[column_post_title].astype(str)

# Functie voor het genereren van een slug
def generate_slug(post_title, focus_keyword, max_retries=3):
    prompt = f"""
    Genereer een SEO-vriendelijke slug gebaseerd op de producttitel '{post_title}' en het focus keyword '{focus_keyword}'. 
    De slug moet:
    - Beginnen met het focus keyword.
    - Bestaan uit maximaal 4-5 woorden en niet meer dan 70 karakters inclusief koppeltekens.
    - Geen speciale tekens of slashes bevatten (alleen letters, cijfers en koppeltekens).
    - Correct Nederlands zijn, zonder afmetingen of irrelevante cijfers.
    """

    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=50  # Beperk tot korte output
            )
            # Verwerk de output naar een geldige slug
            generated_slug = response.choices[0].message.content.strip()
            generated_slug = re.sub(r"[^a-zA-Z0-9\- ]", "", generated_slug)  # Alleen letters, cijfers en koppeltekens
            generated_slug = generated_slug.replace(" ", "-").lower()  # Vervang spaties door koppeltekens en lowercase
            
            # Controleer lengte en beperk tot 4-5 woorden als het te lang is
            if len(generated_slug) > 70:
                words = generated_slug.split("-")
                generated_slug = "-".join(words[:5])

            return generated_slug

        except Exception as e:
            print(f"Error generating slug (attempt {attempt + 1}): {e}")
            time.sleep(1)

    # Als alle pogingen falen, maak een eenvoudige slug
    return f"{focus_keyword}-{post_title[:50]}".lower().replace(" ", "-")

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
    
    post_title = row[column_post_title]
    focus_keyword = row[column_focus_keyword]
    
    # Genereer een nieuwe slug
    new_post_name = generate_slug(post_title, focus_keyword)
    df.at[index, column_post_name] = new_post_name

    # Huidige functies blijven hetzelfde
    new_meta_title = generate_meta_title(focus_keyword)
    df.at[index, column_meta_title] = new_meta_title

    new_meta_description = generate_meta_description(post_title, focus_keyword)
    df.at[index, column_meta_description] = new_meta_description

    # Verwerken van EAN en GTIN
    ean_value = row[column_ean]
    formatted_value = format_gtin_value(ean_value)
    df.at[index, column_gtin] = formatted_value

# Opslaan
output_file = '../herschreven_excel/simpledeal/simpledeal/script2_part1.xlsx'
df.to_excel(output_file, index=False)

print("Verwerking voltooid! Resultaten zijn opgeslagen in:", output_file)