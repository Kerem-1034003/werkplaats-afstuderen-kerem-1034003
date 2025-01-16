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

# Bestanden en paden
excel_file = '../herschreven_excel/simpledeal/split/output_eetkamerstoelen.xlsx'
json_path = '../v10_datamodel_v10_nl.json'
output_file = '../herschreven_excel/simpledeal/homcomimport/updated_eetkamerstoelen.xlsx'

# Mapping-tabel: Excel-categorieën naar JSON-categorieën
category_mapping = {
    "Wonen>Stoelen>Eetkamerstoelen": "Eetkamerstoel",
    "Wonen>Kasten>Nachtkastjes": "Kast",
    "Sport>Fitness & Krachtsport>Aerobic step": "Aerobic stepper",
    "Huisdieren>Honden>Behendigheidspeelgoed": "Speelgoed voor dieren",
    "Wonen>Stoelen>Bureaustoelen": "Bureaustoel"
}

# Functies

def load_json_file(path):
    try:
        with open(path, encoding='utf-8') as file:
            return json.load(file)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"Fout bij het laden van JSON-bestand: {e}")
        exit(1)

def get_required_attributes(category_name, products_data):
    for category, products in products_data.items():
        for product in products:
            if isinstance(product, dict) and product.get('name', '').lower() == category_name.lower():
                return [
                    {
                        'id': attr['id'],
                        'lovId': attr.get('lovId')
                    } 
                    for attr in product.get('attributes', []) 
                    if attr.get('enrichmentLevel') == 1
                ]
    return []

def get_attribute_values(lov_id, products_data):
    if not lov_id:
        return None
    for category, products in products_data.items():
        for product in products:
            if isinstance(product, dict) and product.get('id') == lov_id:
                return product.get('values', [])
    return None

def generate_attribute_value(post_content, attribute_name, allowed_values=None):
    try:
        if allowed_values:
            prompt = f"""
            Op basis van de volgende productomschrijving:
            {post_content}

            Kies een waarde voor het attribuut '{attribute_name}' uit de volgende lijst:
            {', '.join(allowed_values)}

            Geef één waarde, exact zoals in de lijst.
            """
        else:
            prompt = f"""
            Op basis van de volgende productomschrijving:
            {post_content}

            Geef een duidelijke, korte waarde voor het attribuut '{attribute_name}'.
            Ik verwacht alleen de waarde als output en geen tekst ervoor
            Vermijd eenheden zoals 'cm', 'sets' of andere beschrijvingen. Gebruik alleen relevante kernwaarden (zoals een getal, kleur of materiaal).
            """

        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=50,
            temperature=0.7
        )

        value = response.choices[0].message.content.strip()

        # Controleer of de waarde geldig is (indien allowed_values gedefinieerd is)
        if allowed_values:
            if value not in allowed_values:
                print(f"Waarde '{value}' is ongeldig voor '{attribute_name}'. Probeer een fallback te kiezen.")
                # Kies de meest relevante waarde uit allowed_values
                # Hier gebruiken we simpelweg de eerste optie als fallback
                return allowed_values[0] if allowed_values else ""
        
        # Return de gegenereerde waarde als deze geldig is of als er geen restricties zijn
        return value

    except Exception as e:
        print(f"Fout bij het genereren van waarde voor {attribute_name}: {e}")
        return f"Fout bij {attribute_name}"


def map_category(excel_category):
    return category_mapping.get(excel_category, None)

def fill_attribute_values(excel_file, products_data):
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        print(f"Fout bij het lezen van het Excel-bestand: {e}")
        return

    if 'tax:product_cat' not in df.columns:
        print("De kolom 'tax:product_cat' ontbreekt in het Excel-bestand.")
        return

    all_attributes = set()
    for category_name in df['tax:product_cat'].unique():
        json_category = map_category(category_name)
        if not json_category:
            print(f"Geen mapping gevonden voor categorie: {category_name}")
            continue

        required_attributes = get_required_attributes(json_category, products_data)
        all_attributes.update((attr['id'], attr['lovId']) for attr in required_attributes)

    for attr_id, _ in all_attributes:
        if attr_id not in df.columns:
            df[attr_id] = ''

    for index, row in df.iterrows():
        excel_category = row['tax:product_cat']
        json_category = map_category(excel_category)

        if not json_category:
            continue

        required_attributes = get_required_attributes(json_category, products_data)
        post_content = row.get('post_content', '')

        for attribute in required_attributes:
            attr_id = attribute['id']
            lov_id = attribute.get('lovId')

            if pd.isna(row[attr_id]) or row[attr_id] == "":
                allowed_values = get_attribute_values(lov_id, products_data) if lov_id else None
                generated_value = generate_attribute_value(post_content, attr_id, allowed_values)
                df.at[index, attr_id] = generated_value

    try:
        df.to_excel(output_file, index=False)
        print(f"Het bestand is bijgewerkt en opgeslagen als {output_file}")
    except Exception as e:
        print(f"Fout bij het opslaan van het Excel-bestand: {e}")

# Uitvoeren
products_data = load_json_file(json_path)
fill_attribute_values(excel_file, products_data)
