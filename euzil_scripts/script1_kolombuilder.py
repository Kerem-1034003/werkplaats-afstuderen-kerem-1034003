import pandas as pd

# Functie om de Excel-kolommen aan te passen
def modify_excel_columns(input_file, output_file, additional_file):
    # Laad de Excel-bestanden in DataFrames
    df = pd.read_excel(input_file)
    additional_df = pd.read_excel(additional_file)

    # Kolommen verwijderen
    columns_to_remove = ['Title', 'MSRP', 'Material', 'Brand']
    df.drop(columns=columns_to_remove, inplace=True, errors='ignore')

    # Kolommen hernoemen
    columns_to_rename = {
        'SKU': 'sku',
        'EAN': 'meta:_alg_ean',
        'Name': 'post_title',
        'Category': 'tax:product_cat'
    }
    df.rename(columns=columns_to_rename, inplace=True)

    # Kolommen toevoegen
    additional_columns = [
        'post_name',
        'post_content',
        'post_excerpt',
        'meta:_yoast_wpseo_focuskw',
        'meta:_yoast_wpseo_metadesc',
        'meta:_yoast_wpseo_title',
        'meta:wpseo_global_identifier_values',
        'images'
    ]
    for col in additional_columns:
        df[col] = ''  # Voeg lege kolommen toe

    # Kolom en waarden uit additional_file toevoegen op basis van SKU
    additional_column = '（1-4PCS）'  # Kolomnaam in het tweede Excel-bestand
    if additional_column in additional_df.columns:
        # Controleer of de SKU-kolom overeenkomt in beide bestanden
        additional_df.rename(columns={'SKU': 'sku'}, inplace=True)
        df = pd.merge(df, additional_df[['sku', additional_column]], on='sku', how='left')
    else:
        print(f"Kolom '{additional_column}' niet gevonden in {additional_file}")

    # Verwerkte DataFrame opslaan naar een nieuw Excel-bestand
    df.to_excel(output_file, index=False)
    print(f"Aangepaste Excel opgeslagen in: {output_file}")

# Input- en outputbestanden
input_file = '../excel/simpledeal/euzil/Product Informaton 20241205.xlsx'  # Vervang door je invoerbestand
additional_file = '../excel/simpledeal/euzil/wholesale price ZAZA-2024.10.10.xlsx'  # Bestand met aanvullende gegevens
output_file = '../excel/simpledeal/euzil/euzil.xlsx'  # Vervang door je gewenste uitvoerbestand

# Voer de functie uit
modify_excel_columns(input_file, output_file, additional_file)
