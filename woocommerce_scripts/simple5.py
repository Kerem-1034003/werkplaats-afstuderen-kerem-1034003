import pandas as pd
import os

# Bestand inleiden
df = pd.read_excel('../herschreven_excel/simpledeal/simpledeal/script4_part1.xlsx')

# Zorg ervoor dat de uitvoermap bestaat
output_folder = '../herschreven_excel/simpledeal/simpledealsplit/'
os.makedirs(output_folder, exist_ok=True)

# Splitsen op basis van de 'tax:product_cat' kolom
categories = df['tax:product_cat'].unique()

# Voor elke unieke categorie, maak een nieuw bestand
for category in categories:
    # Verkrijg de laatste categorie (na de laatste '>' of als er geen '>' is, de hele string)
    last_category = category.split('>')[-1]
    
    # Vervang ongeldige tekens in de bestandsnaam, zoals '>', met '_'
    last_category = last_category.replace('>', '_').replace('/', '_').replace('\\', '_')

    # Filter de producten per categorie
    category_df = df[df['tax:product_cat'] == category]
    
    # Bestandsnaam op basis van de laatste categorie
    output_file = f'{output_folder}output_{last_category}.xlsx'
    
    # Sla het gefilterde bestand op
    category_df.to_excel(output_file, index=False)

    print(f'Bestand voor categorie "{last_category}" opgeslagen als: {output_file}')
