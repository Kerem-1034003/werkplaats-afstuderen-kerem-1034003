import pandas as pd

# Functie om de Excel-kolommen aan te passen
def modify_excel_columns(input_file, output_file):
    # Laad de Excel in een DataFrame
    df = pd.read_excel(input_file)

    # Kolommen verwijderen
    columns_to_remove = ['Color', 'Material', 'Brand']
    df.drop(columns=columns_to_remove, inplace=True, errors='ignore')

    # Kolommen hernoemen
    columns_to_rename = {
        'SKU': 'sku',
        'EAN': 'meta:_alg_ean',
        'Name': 'post_title',
        'Description': 'post_content',
        'Category' :  'tax:product_cat'
    }
    df.rename(columns=columns_to_rename, inplace=True)

    # Kolommen toevoegen
    additional_columns = [
        'post_name',
        'post_excerpt',
        'meta:_yoast_wpseo_focuskw',
        'meta:_yoast_wpseo_metadesc',
        'meta:_yoast_wpseo_title',
        'meta:wpseo_global_identifier_values',
        'images'
    ]
    for col in additional_columns:
        df[col] = ''  # Voeg lege kolommen toe

    # Verwerkte DataFrame opslaan naar een nieuw Excel-bestand
    df.to_excel(output_file, index=False)
    print(f"Aangepaste Excel opgeslagen in: {output_file}")

# Input- en outputbestanden
input_file = '../excel/simpledeal/homcom/homcom.xlsx'  # Vervang door je invoerbestand
output_file = '../excel/simpledeal/homcom/homcom1.xlsx'  # Vervang door je gewenste uitvoerbestand

# Voer de functie uit
modify_excel_columns(input_file, output_file)
