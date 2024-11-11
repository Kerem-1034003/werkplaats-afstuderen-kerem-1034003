import os
from dotenv import load_dotenv
import pandas as pd
from openai import OpenAI
import time

# Laad de OpenAI API key
load_dotenv()
openai_api_key = os.getenv('OPENAI_API_KEY2')

# Initialiseer de OpenAI client
client = OpenAI(api_key=openai_api_key)

# Bedrijfsnaam toevoegen
company_name = "Simpledeal"

# Lees de Simpledeal Excel-bestand
df = pd.read_excel('excel/simpledeal/simpledeal-cat.xlsx')

# Zorg dat de `meta:_yoast_data` kolom tekst kan opslaan
df['meta:_yoast_data'] = df['meta:_yoast_data'].astype(str)
df['description'] = df['description'].astype(str)

predefined_translations = {
    "Beautycases": "Toilettassen",
    "SackCarrow": "Steekwagen",
}

# Functie voor vertalen en corrigeren van `name`
def translate_and_correct_category(category_name):
    # Check of er een vooraf gedefinieerde vertaling is
    if category_name in predefined_translations:
        return predefined_translations[category_name]
    
    # Specifieke vertaling van & behouden
    prompt = f"""
    Als de productcategorie '{category_name}' niet in het Nederlands is, vertaal deze dan naar correct Nederlands.
    Als de naam al in het Nederlands is maar typfouten of onjuiste samenstellingen bevat, verbeter deze dan.
    Als er een '&' (ampersand) in de naam voorkomt, behoud deze dan en vervang deze niet door 'en'.
    Geef alleen de naam als output zonder extra uitleg of symbolen.
    """

    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=20
    )

    # Return de vertaalde of gecorrigeerde naam
    return response.choices[0].message.content.strip()

# Functie voor beschrijving van de categorie
def generate_category_description(category_name):
    prompt = f"""
    Schrijf een uitgebreide en informatieve beschrijving van minimaal 500 tot maximaal 550 woorden voor de productcategorie '{category_name}'.
    Zorg dat het de producten in deze categorie beschrijft en gericht is op het aantrekken van potentiële klanten.
    """
    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=700
    )
    return response.choices[0].message.content.strip()  # Correcte manier om de response te krijgen

# Functie om `meta:_yoast_data` op te bouwen met bedrijfsnaam-aanvulling en permalink
def build_meta_yoast_data(focus_keyword):
    # Genereer de meta title
    meta_title_prompt = f"""
    Maak een aantrekkelijke en SEO-geoptimaliseerde meta title beginnend met '{focus_keyword}'.
    De titel moet bestaan uit 5-6 woorden, en mag niet langer zijn dan 60 karakters inclusief spaties.
    Zorg ervoor dat de titel niet wordt afgebroken en aantrekkelijk is voor de lezer.
    """
    meta_title_response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": meta_title_prompt}],
        max_tokens=50
    )
    meta_title = meta_title_response.choices[0].message.content.strip()  # Correcte manier om de response te krijgen

    # Controleer lengte van meta title en voeg bedrijfsnaam toe als deze korter is dan 47 karakters
    if len(meta_title) < 47:
        meta_title = f"{meta_title} | {company_name}"

    # Genereer de meta description
    meta_desc_prompt = f"""
    Schrijf een korte, informatieve meta description van maximaal 150 karakters beginnend met '{focus_keyword}'.
    """
    meta_desc_response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": meta_desc_prompt}],
        max_tokens=150
    )
    meta_description = meta_desc_response.choices[0].message.content.strip()  # Correcte manier om de response te krijgen

    # Genereer Yoast data in het gewenste formaat
    yoast_data = f"""a:7:{{s:21:"wpseo_keywordsynonyms";s:4:"[""]";s:11:"wpseo_title";s:{len(meta_title)}:"{meta_title}";s:10:"wpseo_desc";s:{len(meta_description)}:"{meta_description}";s:13:"wpseo_focuskw";s:{len(focus_keyword)}:"{focus_keyword}";s:13:"wpseo_linkdex";s:;s:19:"wpseo_content_score";s:1:"0";}}"""
    return yoast_data

# Verwerk de rijen in de DataFrame
for idx, row in df.iterrows():
    print(f"Processing row {idx + 1} of {len(df)}")
    
    # Stap 1: Vertaal en corrigeer de categorienaam
    category_name = row['name']
    corrected_name = translate_and_correct_category(category_name)
    df.at[idx, 'name'] = corrected_name

    # Vul de slug-kolom in met de nieuwe naam, maar dan in kleine letters
    df.at[idx, 'slug'] = corrected_name.lower()

    # Genereer beschrijving voor de categorie
    category_description = generate_category_description(corrected_name)
    df.at[idx, 'description'] = category_description

    # Genereer meta:_yoast_data
    focus_keyword = corrected_name  # Gebruik de herschreven categorienaam als focus keyword
    meta_yoast_data = build_meta_yoast_data(focus_keyword)
    df.at[idx, 'meta:_yoast_data'] = meta_yoast_data

# Opslaan naar een nieuw Excel-bestand
df.to_excel('herschreven_excel/simpledeal/simpledealcat.xlsx', index=False)
print("Verwerking voltooid en opgeslagen")
