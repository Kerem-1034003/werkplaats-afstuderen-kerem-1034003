import os
import re
from dotenv import load_dotenv
import pandas as pd
from openai import OpenAI
import time

# Laad de API-sleutel
load_dotenv()
openai_api_key = os.getenv('OPENAI_API_KEY2')

# Initialiseer de OpenAI client
client = OpenAI(api_key=openai_api_key)

# Laad de DataFrame
df = pd.read_excel('../herschreven_excel/simpledeal/simpledeal/script3_part1.xlsx')

column_images = 'images'
column_focus_keyword = 'meta:_yoast_wpseo_focuskw'

# Functie om alt-tekst te genereren en toe te voegen
def generate_and_add_alt_text(df, images_column='images', focus_keyword_column='meta:_yoast_wpseo_focuskw', max_retries=3):
    for idx, row in df.iterrows():
        print(f"Bezig met rij {idx + 1} van {len(df)}")  # Print huidige rij
        images_data = row[images_column]
        focus_keyword = row[focus_keyword_column]

        # Splits de structuur van de images op de pipe '|' separator
        image_entries = images_data.split(' | ')

        updated_image_entries = []

        for entry in image_entries:
            # Extracteer de afbeelding link
            match = re.search(r'(https?://\S+)', entry)
            if match:
                image_url = match.group(1)
                
                # Genereer alt-tekst met behulp van OpenAI
                prompt = f"""
                Genereer een unieke en beschrijvende alt-tekst van ongeveer 75 karakters voor een productafbeelding met de volgende link: '{image_url}'. 
                Begin met het focus keyword: '{focus_keyword}'. Houd de beschrijving relevant en helder, en in correct Nederlands. Gebruik geen woorden zoals "Alt-tekst:" in de alt-tekst.Voeg ook een vleugje variatie toe.
                """

                alt_text = ""
                for attempt in range(max_retries):
                    try:
                        response = client.chat.completions.create(
                            model="gpt-3.5-turbo",
                            messages=[{"role": "user", "content": prompt}],
                            max_tokens=60,
                            temperature=0.7
                        )

                        alt_text = response.choices[0].message.content.strip()
                        break
                    except Exception as e:
                        print(f"Error generating alt text for image {image_url} on attempt {attempt + 1}: {e}")
                        time.sleep(2)

                # Voeg de gegenereerde alt-tekst toe aan de entry
                entry = re.sub(r'alt :\s*!', f'alt : {alt_text} !', entry)

            updated_image_entries.append(entry)

        # Combineer de aangepaste entries weer in de originele structuur
        df.at[idx, images_column] = ' | '.join(updated_image_entries)

    return df

# Verwerk de DataFrame
df = generate_and_add_alt_text(
    df,
    images_column=column_images,
    focus_keyword_column=column_focus_keyword
)

output_file = '..herschreven_excel/simpledeal/simpledeal/script4_part1.xlsx'
df.to_excel(output_file, index=False)

print("Verwerking voltooid! Resultaten zijn opgeslagen in:", output_file)
