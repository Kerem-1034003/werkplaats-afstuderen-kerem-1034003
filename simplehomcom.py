import os
from dotenv import load_dotenv
import pandas as pd
from openai import OpenAI
import time

# Laad de OpenAI API key
load_dotenv()
openai_api_key = os.getenv('OPENAI_API_KEY')

# Initialiseer de OpenAI client
client = OpenAI(api_key=openai_api_key)

# Bedrijfsnaam toevoegen
company_name = "Simpledeal"

# Lees de Simpledeal Excel-bestand
df = pd.read_excel('excel/simpledeal/Map1.xlsx')

# Zorg dat de `meta:_yoast_data` kolom tekst kan opslaan
df['meta:_yoast_data'] = df['meta:_yoast_data'].astype(str)
df['description'] = df['description'].astype(str)

predefined_translations = {
    "Beautycases": "Toilettassen",
    "SackCarrow": "Steekwagen",
}

# Functie voor vertalen en corrigeren van `name`
def translate_and_correct_category(category_name):
    if category_name in predefined_translations:
        return predefined_translations[category_name]
    
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

    return response.choices[0].message.content.strip()

def generate_category_description(category_name):
    prompt = f"""
    Schrijf een uitgebreide en informatieve beschrijving van minimaal 500 woorden voor de productcategorie '{category_name}'. De beschrijving moet voldoen aan de volgende voorwaarden:
    1. Het focus keyword '{category_name}' moet minimaal vijf keer voorkomen in de tekst.
    2. Gebruik HTML-formattering met korte paragrafen en koppen. Hierbij verwacht ik kopteksten 
    3. De tekst moet gericht zijn op het aantrekken van potentiële klanten, waarbij de producten in deze categorie worden beschreven.
    4. Gebruik korte en duidelijke paragrafen (maximaal 3-4 zinnen per alinea).
    5. Sluit de beschrijving goed af, bij voorkeur met een call-to-action (bijvoorbeeld een uitnodiging om producten te bekijken).
    """

    # Maak een lange prompt voor OpenAI en stel een hogere max_tokens waarde in
    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=1500  # Zorg ervoor dat we voldoende tokens krijgen voor een lange beschrijving
    )

    return response.choices[0].message.content.strip()


def build_meta_yoast_data(focus_keyword):
    # Genereer de meta title
    meta_title_prompt = f"""
    Maak een aantrekkelijke en SEO-geoptimaliseerde meta title beginnend met '{focus_keyword}'.
    De titel moet bestaan uit 5-6 woorden en mag niet langer zijn dan 60 karakters inclusief spaties.
    Zorg ervoor dat de titel niet wordt afgebroken en aantrekkelijk is voor de lezer.
    """
    meta_title_response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": meta_title_prompt}],
        max_tokens=50
    )
    meta_title = meta_title_response.choices[0].message.content.strip()

    # Voeg bedrijfsnaam toe indien titel korter is dan 47 karakters
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
    meta_description = meta_desc_response.choices[0].message.content.strip()

    # Bepaal de waardes en lengtes voor serialisatie
    focuskw_length = len(focus_keyword)
    title_length = len(meta_title)
    desc_length = len(meta_description)
    linkdex_value = "53"  # Voorbeeldscore, kan ook dynamisch worden berekend of opgehaald uit je data
    content_score = "0"   # Standaardwaarde content score

    # Bouw de geserialiseerde Yoast-data string (op basis van het voorbeeld)
    yoast_data = f"""a:5:{{s:11:"wpseo_title";s:{title_length}:"{meta_title}";s:10:"wpseo_desc";s:{desc_length}:"{meta_description}";s:13:"wpseo_focuskw";s:{focuskw_length}:"{focus_keyword}";s:13:"wpseo_linkdex";s:2:"{linkdex_value}";s:19:"wpseo_content_score";s:1:"{content_score}";}}"""

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
df.to_excel('herschreven_excel/simpledeal/simpledealcat1.2.xlsx', index=False)
print("Verwerking voltooid en opgeslagen")
