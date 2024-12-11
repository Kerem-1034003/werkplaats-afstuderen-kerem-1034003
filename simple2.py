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
df = pd.read_excel('herschreven_excel/simpledeal/homcomproducten_2.xlsx')

column_focus_keyword = 'meta:_yoast_wpseo_focuskw'
column_images = 'images'

# Functie voor het toevoegen van alt-tekst aan afbeeldingen
def add_alt_text_to_images(df, images_column='images', focus_keyword_column='meta:rank_math_focus_keyword', max_retries=3):
    for idx, row in df.iterrows():
        if pd.notna(row[images_column]):
            images = row[images_column].split(' | ')
        if images:
            prompt = f"""
            Genereer een korte en beschrijvende alt-tekst van ongeveer 75 karakters voor een productafbeelding, beginnend met het focus keyword: '{row[focus_keyword_column]}'. 
            Houd de beschrijving relevant en helder, en in correct Nederlands.
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
                parts = images[0].split(' ! ')
                new_parts = []
                for part in parts:
                    if part.startswith('alt :'):
                        new_parts.append(f"alt : {alt_text}")
                    else:
                        new_parts.append(part)
                images[0] = ' ! '.join(new_parts)
                break

            except Exception as e:
                print(f"Error generating alt text for row {idx} on attempt {attempt + 1}: {e}")
                time.sleep(2)

        df.at[idx, images_column] = ' | '.join(images)

    return df


# Verwerking van de DataFrame
for index, row in df.iterrows():
    print(f"Processing row {index + 1} of {len(df)}")

df = add_alt_text_to_images(
        df, 
        images_column='images', 
        focus_keyword_column=column_focus_keyword
)

output_file = 'herschreven_excel/simpledeal/homcomproducten_2.1.xlsx'
df.to_excel(output_file, index=False)

print("Verwerking voltooid! Resultaten zijn opgeslagen in:", output_file)