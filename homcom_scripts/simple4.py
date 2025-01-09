import os
import re
from dotenv import load_dotenv
import pandas as pd
from openai import OpenAI
import time
import json

load_dotenv()

openai_api_key = os.getenv('OPENAI_API_KEY2')

# Initialiseer de OpenAI client
client = OpenAI(api_key=openai_api_key)

# Laad de DataFrame
df = pd.read_excel('../herschreven_excel/simpledeal/homcom/homcom2.3.xlsx')

column_focus_keyword = 'meta:_yoast_wpseo_focuskw'
column_images = 'images'
column_image_1_link = 'Image 1 Link'
column_image_additional_links = 'Image Additional Links'

df[column_images] = df[column_images].astype(str)

# Functie voor het toevoegen van alt-tekst aan afbeeldingen
def add_alt_text_to_images(df, images_column='images', focus_keyword_column='meta:_yoast_wpseo_focuskw', max_retries=3):
    for idx, row in df.iterrows():
        images = []

        # Samengevoegde images uit beide kolommen
        if pd.notna(row[column_image_1_link]):
            images.append(row[column_image_1_link])
        if pd.notna(row[column_image_additional_links]):
            # Vervang komma's door pipes en split dan de tekst
            additional_images = row[column_image_additional_links].replace(',', ' | ').split(' | ')
            images.extend(additional_images)

        # De gewenste afbeelding opbouwen
        image_parts = []
        for image in images:
            image_parts.append(f"{image} ! alt :  ! title : {image.split('/')[-1]} ! desc :  ! caption : ")

        # Zet de samengestelde images in de 'images' kolom
        df.at[idx, images_column] = ' | '.join(image_parts)

        # Genereer alt-tekst voor elke afbeelding
        for i, image in enumerate(images):
            prompt = f"""
            Genereer een unieke en beschrijvende alt-tekst van ongeveer 75 karakters voor een productafbeelding met de volgende link: '{image}'. 
            Begin met het focus keyword: '{row[focus_keyword_column]}'. Houd de beschrijving relevant en helder, en in correct Nederlands. Gebruik geen woorden zoals "Alt-tekst:" in de alt-tekst.
            """
            for attempt in range(max_retries):
                try:
                    response = client.chat.completions.create(
                        model="gpt-3.5-turbo",
                        messages=[{"role": "user", "content": prompt}],
                        max_tokens=60,
                        temperature=0.7
                    )

                    alt_text = response.choices[0].message.content.strip()
                    image_parts[i] = image_parts[i].replace('alt :  ', f'alt : {alt_text} ')
                    break

                except Exception as e:
                    print(f"Error generating alt text for image {image} on attempt {attempt + 1}: {e}")
                    time.sleep(2)

        # Zet de nieuwe afbeeldingstekst (met alt-teksten) in de 'images' kolom
        df.at[idx, images_column] = ' | '.join(image_parts)

    return df

# Verwerking van de DataFrame
for index, row in df.iterrows():
    print(f"Processing row {index + 1} of {len(df)}")

df = add_alt_text_to_images(
    df, 
    images_column='images', 
    focus_keyword_column=column_focus_keyword
)

output_file = '../herschreven_excel/simpledeal/homcom/homcom_simple2.xlsx'
df.to_excel(output_file, index=False)

print("Verwerking voltooid! Resultaten zijn opgeslagen in:", output_file)
