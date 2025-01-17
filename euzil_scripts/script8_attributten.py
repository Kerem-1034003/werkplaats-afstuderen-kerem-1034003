import pandas as pd
import json
import os
from openai import OpenAI
from dotenv import load_dotenv
import re

load_dotenv()

openai_api_key = os.getenv('OPENAI_API_KEY')

# Initialiseer de OpenAI client
client = OpenAI(api_key=openai_api_key)

# Gebruik de categorie en het Excel-bestand om de attributen toe te voegen en te vullen
excel_file = '../herschreven_excel/simpledeal/euzilsplit/output_Badkamerkast.xlsx'  # Het bestand van de producten

# Laad je JSON-bestand met producten
with open('../v10_datamodel_v10_nl.json', encoding='utf-8') as json_file:
    products_data = json.load(json_file)

# Mapping-tabel: Excel-categorieën naar JSON-categorieën
category_mapping = {
    "Wonen>Stoelen>Eetkamerstoelen": "Eetkamerstoel",
    "Wonen>Kasten>Nachtkastjes": "Kast",
    "Sport>Fitness & Krachtsport>Aerobic step": "Aerobic stepper",
    "Huisdieren>Honden>Behendigheidspeelgoed": "Speelgoed voor dieren",
    "Wonen>Stoelen>Bureaustoelen": "Bureaustoel",
    "Wonen>Kasten>Badkamerkast" : "Badkamerkast"

    # Voeg meer categorieën toe indien nodig
}

# Functie om de vereiste attributen voor een productcategorie op te halen
def get_required_attributes(category_name):
    for category, products in products_data.items():
        for product in products:
            if isinstance(product, dict) and 'name' in product and product['name'].lower() == category_name.lower():
                required_attributes = [
                    attribute['id'] for attribute in product.get('attributes', [])
                    if attribute.get('enrichmentLevel') == 1
                ]
                return required_attributes
    return None  # Return None als geen match gevonden

# Functie om de juiste JSON-categorie te vinden op basis van de mapping
def map_category(excel_category):
    return category_mapping.get(excel_category, None)  # Retourneer None als er geen mapping is

def generate_attribute_value(post_content, attribute_name):
    try:
        prompt = f"""
        Op basis van de volgende productomschrijving:
        {post_content}
        
        Geef een waarde voor het attribuut '{attribute_name}'. Zorg dat de waarde kort, relevant en duidelijk is, zonder de kolomnaam of extra tekens zoals dubbele punt (:). Als het antwoord 'Ja' of 'Nee' is, gebruik dan 'Yes' of 'No' in het Engels. Zorg ervoor dat eenheden zoals 'cm' of 'sets' worden weggelaten en alleen de getallen behouden blijven.
        """
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=50,
            temperature=0.7
        )

        # Haal de juiste waarde uit de API-respons
        generated_value =  response.choices[0].message.content.strip().replace('"', '').replace("'", "")

        # Converteer 'Ja'/'Nee' naar 'Yes'/'No' als dat nodig is
        if generated_value.lower() == 'ja':
            generated_value = 'Yes'
        elif generated_value.lower() == 'nee':
            generated_value = 'No'

        # Verwijder eenheden zoals 'cm' of 'sets' en behoud alleen getallen
        generated_value = re.sub(r'\s*(cm|sets)\s*', '', generated_value)

        return generated_value   
    except Exception as e:
        print(f"Fout bij het genereren van waarde voor {attribute_name}: {e}")
        return ""


# Functie om het Excel-bestand bij te werken met de vereiste attributen als kolommen
def fill_attribute_values(excel_file):
    # Lees de Excel-bestand
    df = pd.read_excel(excel_file)

    # Controleer of de categorie-kolom aanwezig is
    if 'tax:product_cat' not in df.columns:
        print("De kolom 'tax:product_cat' ontbreekt in het Excel-bestand.")
        return

    # Voeg kolommen toe voor de vereiste attributen, indien nog niet aanwezig
    for index, row in df.iterrows():
        excel_category = row['tax:product_cat']
        json_category = map_category(excel_category)

        if not json_category:
            print(f"Geen mapping gevonden voor categorie: {excel_category}")
            continue

        # Haal de vereiste attributen op voor de gemapte categorie
        required_attributes = get_required_attributes(json_category)

        if not required_attributes:
            print(f"Geen vereiste attributen gevonden voor JSON-categorie: {json_category}")
            continue

        # Voeg kolommen toe voor de vereiste attributen, indien nog niet aanwezig
        for attribute in required_attributes:
            if attribute not in df.columns:
                df[attribute] = ''  # Voeg lege kolommen toe

    # Loop door de rijen om attributen te verwerken en waarden te genereren
    for index, row in df.iterrows():
        excel_category = row['tax:product_cat']
        json_category = map_category(excel_category)

        if not json_category:
            print(f"Geen mapping gevonden voor categorie: {excel_category}")
            continue

        # Haal de vereiste attributen op voor de gemapte categorie
        required_attributes = get_required_attributes(json_category)

        if not required_attributes:
            print(f"Geen vereiste attributen gevonden voor JSON-categorie: {json_category}")
            continue

        # Gebruik de productomschrijving (post_content) om de waarden te genereren
        post_content = row['post_content'] if 'post_content' in row else ''

        for attribute in required_attributes:
            if attribute in df.columns:
                # Alleen invullen als het veld leeg is
                if pd.isna(row[attribute]) or row[attribute] == "":
                    generated_value = generate_attribute_value(post_content, attribute)
                    print(f"Generated value for {attribute}: {generated_value}")  # Debugging
                    df.at[index, attribute] = generated_value

    # Specificeer de directory voor het uitvoerbestand
    output_file = '../herschreven_excel/simpledeal/euzilimport/updated_badkamerkast.xlsx'

    # Bewaar de gewijzigde DataFrame naar het opgegeven bestand
    df.to_excel(output_file, index=False)

    print(f"Het bestand is bijgewerkt met gegenereerde waarden en opgeslagen als {output_file}")

    return output_file


# Test de functie om het Excel-bestand bij te werken
updated_file = fill_attribute_values(excel_file)

if updated_file:
    print(f"Het bestand is bijgewerkt en opgeslagen als {updated_file}")
