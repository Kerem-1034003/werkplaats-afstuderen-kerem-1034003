import pandas as pd
import json
import os

# Gebruik de categorie en het Excel-bestand om de attributen toe te voegen en te vullen
excel_file = '../herschreven_excel/simpledeal/homcomsplit/output_Eetkamerstoelen.xlsx'  # Het bestand van de producten

# Laad je JSON-bestand met producten
with open('../v10_datamodel_v10_nl.json', encoding='utf-8') as json_file:
    products_data = json.load(json_file)

# Voorbeeld van een mapping-tabel: Excel-categorieën naar JSON-categorieën
category_mapping = {
    "Wonen>Stoelen>Eetkamerstoelen": "Eetkamerstoel",
    "Wonen>Kasten>Nachtkastjes": "Kast",
    # Voeg hier meer categorieën toe als nodig
}

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

# Functie om de juiste JSON-categorie te vinden op basis van de mapping
def map_category(excel_category):
    return category_mapping.get(excel_category, None)  # Retourneer None als er geen mapping is

# Functie om het Excel-bestand bij te werken met de vereiste attributen als kolommen
def update_excel_with_attributes(excel_file):
    # Lees de Excel-bestand
    df = pd.read_excel(excel_file)

    # Controleer of de categorie-kolom aanwezig is
    if 'tax:product_cat' not in df.columns:
        print("De kolom 'tax:product_cat' ontbreekt in het Excel-bestand.")
        return

    # Loop door de rijen om de categorieën te verwerken
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

    # Specificeer de directory voor het uitvoerbestand
    output_file = '../herschreven_excel/simpledeal/homcom/updated_eetkamerstoelen.xlsx'

    # Bewaar de gewijzigde DataFrame naar het opgegeven bestand
    df.to_excel(output_file, index=False)

    print(f"Het bestand is bijgewerkt en opgeslagen als {output_file}")

    return output_file

# Test de functie om het Excel-bestand bij te werken
updated_file = update_excel_with_attributes(excel_file)

if updated_file:
    print(f"Het bestand is bijgewerkt en opgeslagen als {updated_file}")
