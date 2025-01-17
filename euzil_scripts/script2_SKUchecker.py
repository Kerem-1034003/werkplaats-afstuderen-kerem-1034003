import pandas as pd

# Functie om nieuwe producten te filteren op basis van SKU
def filter_new_products(current_products_file, supplier_products_file, output_file, sku_column='sku'):
    # Laad de huidige producten en leverancier producten in DataFrames
    current_df = pd.read_excel(current_products_file)
    supplier_df = pd.read_excel(supplier_products_file)

    # Zorg ervoor dat de SKU-kolommen als strings worden behandeld
    current_df[sku_column] = current_df[sku_column].astype(str)
    supplier_df[sku_column] = supplier_df[sku_column].astype(str)

    # Identificeer de SKU's die niet in de huidige producten zitten
    current_skus = set(current_df[sku_column])
    new_products = supplier_df[~supplier_df[sku_column].isin(current_skus)]

    # Sla de nieuwe producten op in een nieuw Excel-bestand
    new_products.to_excel(output_file, index=False)
    print(f"Nieuwe producten zijn opgeslagen in: {output_file}")

# Bestanden en kolomnaam instellen
current_products_file = '../excel/simpledeal/euzil/euzil.xlsx' # Vervang door je huidige productbestand
supplier_products_file = '../excel/simpledeal/woocommerce/euzil_producten.xlsx'  # Vervang door het bestand van de leverancier
output_file = '../excel/simpledeal/euzil/oude_euzilproducten.xlsx'  # Uitvoerbestand voor nieuwe producten
sku_column = 'sku'  # Pas aan indien nodig

# Filter nieuwe producten
filter_new_products(current_products_file, supplier_products_file, output_file, sku_column)
