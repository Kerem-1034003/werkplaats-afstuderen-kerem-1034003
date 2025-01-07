import pandas as pd
import json
import openai
import os

# Gebruik de categorie en het Excel-bestand om de attributen toe te voegen en te vullen
excel_file = '../herschreven_excel/simpledeal/homcomsplit/output_Eetkamerstoelen.xlsx'  # Het bestand van de producten

# Laad je JSON-bestand met producten
with open('../v10_datamodel_v10_nl.json', encoding='utf-8') as json_file:
    products_data = json.load(json_file)

# Functie om de vereiste attributen voor een productcategorie op te halen
def get_required_attributes(category_name):
    for category, products in products_data.items():
        for product in products:
            if isinstance(product, dict) and 'name' in product and product['name'].lower() == category_name.lower():
                # Filter de attributen met enrichmentLevel == 1
                required_attributes = [
                    attribute['id'] for attribute in product.get('attributes', [])
                    if attribute.get('enrichmentLevel') == 1
                ]
                return required_attributes
    return None  # Return None als geen match gevonden

# Functie om het Excel-bestand bij te werken met de vereiste attributen als kolommen
def update_excel_with_attributes(excel_file, category_name):
    # Lees de Excel-bestand
    df = pd.read_excel(excel_file)

    # Verkrijg de vereiste attributen voor de opgegeven categorie
    required_attributes = get_required_attributes(category_name)

    if not required_attributes:
        print(f"Geen vereiste attributen gevonden voor categorie: {category_name}")
        return

    # Maak kolommen aan voor de vereiste attributen in de DataFrame
    for attribute in required_attributes:
        if attribute not in df.columns:
            df[attribute] = ''  # Voeg lege kolommen toe

    # Specificeer de directory voor het uitvoerbestand
    output_file = '../herschreven_excel/simpledeal/homcom/updated_eetkamerstoelen.xlsx'

    # Bewaar de gewijzigde DataFrame naar het opgegeven bestand
    df.to_excel(output_file, index=False)

    print(f"Het bestand is bijgewerkt en opgeslagen als {output_file}")

    return output_file

# Test de functie om het Excel-bestand bij te werken
category_name = 'Eetkamerstoel'
updated_file = update_excel_with_attributes(excel_file, category_name)

if updated_file:
    print(f"Het bestand is bijgewerkt en opgeslagen als {updated_file}")
