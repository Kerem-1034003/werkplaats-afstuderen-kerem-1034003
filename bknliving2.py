import os
from dotenv import load_dotenv
import pandas as pd
from openai import OpenAI
import time

# Laad de OpenAI API key
load_dotenv()
openai_api_key = os.getenv('OPENAI_API_KEY')

# Initialiseer de OpenAI client
client = OpenAI(api_key=openai_api_key)

df = pd.read_excel('herschreven_excel/bkn-living/bknliving1.1.xlsx')

# Definieer de kolomnamen
column_post_name = 'post_name'
column_focus_keyword = 'meta:rank_math_focus_keyword'
column_alg_ean = 'meta:_alg_ean'
column_global_unique_id = 'meta:_global_unique_id'
column_images = 'images'

# Functie voor het herschrijven van de URL (post_name)
def rewrite_post_name(focus_keyword, post_name):
    try:
        prompt = f"""
        Herschrijf de URL '{post_name}' zodat deze begint met het focus keyword '{focus_keyword}' 
        en maximaal 70 tekens bevat inclusief spaties.
        """
        
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=70
        )
        
        new_post_name = response.choices[0].message.content.strip().replace('"', '').replace("'", "")
        return new_post_name
    except Exception as e:
        print(f"Error rewriting post name: {e}")
        return post_name

# Functie om EAN naar global_unique_id te kopiëren
def copy_ean_to_global_unique_id(df, ean_column='meta:_alg_ean', global_unique_id_column='meta:_global_unique_id'):
    for idx, row in df.iterrows():
        ean_value = row[ean_column]
        if pd.notna(ean_value):  # Controleer of er een waarde aanwezig is in de kolom 'meta:_alg_ean'
            df.at[idx, global_unique_id_column] = ean_value  # Kopieer de waarde naar 'meta:_global_unique_id'
    return df

# Functie voor het toevoegen van alt-tekst aan afbeeldingen
def add_alt_text_to_images(df, images_column='images', focus_keyword_column='meta:rank_math_focus_keyword', max_retries=3):
    for idx, row in df.iterrows():
        if pd.notna(row[images_column]):
            images = row[images_column].split(' | ')  # Splits de afbeeldingen in de lijst
            if images:
                prompt = f"""
                Genereer een korte en beschrijvende alt-tekst van ongeveer 75 karakters voor een productafbeelding.
                Gebruik het volgende focus keyword: '{row[focus_keyword_column]}'. 
                Houd de beschrijving relevant en helder, en in correct Nederlands.
                """
                
                # Ingebouwd retry mechanisme voor timeouts en rate limits
                for attempt in range(max_retries):
                    try:
                        response = client.chat.completions.create(
                            model="gpt-3.5-turbo",
                            messages=[{"role": "user", "content": prompt}],
                            max_tokens=60,  # Houd het kort voor snellere response
                            temperature=0.7
                        )
                        
                        alt_text = response.choices[0].message.content.strip()
                        
                        parts = images[0].split(' ! ')
                        new_parts = []
                        for part in parts:
                            if part.startswith('alt :'):
                                new_parts.append(f"alt : {alt_text}")  # Vervang de alt-tekst
                            else:
                                new_parts.append(part)
                        images[0] = ' ! '.join(new_parts)  # Werk de eerste afbeelding bij
                        
                        break
                    
                    except client.error.APIError as e:
                        print(f"APIError on attempt {attempt + 1} for row {idx}: {e}")
                        time.sleep(2)  # Wacht 2 seconden voor nieuwe poging
                    except client.error.RateLimitError:
                        print(f"Rate limit hit on attempt {attempt + 1} for row {idx}. Retrying in 10 seconds...")
                        time.sleep(10)  # Wacht langer bij rate limit error
                    except Exception as e:
                        print(f"Onverwachte fout bij rij {idx}: {e}")
                        break  # Breek uit de loop bij andere fouten

                df.at[idx, images_column] = ' | '.join(images)

    return df

# Itereer door elke rij in de DataFrame voor herschrijven van post_name en alt-tekst
for index, row in df.iterrows():
    print(f"Processing row {index + 1} of {len(df)}")
    
    focus_keyword = row[column_focus_keyword]
    post_name = row[column_post_name]
    
    # URL herschrijven
    new_post_name = rewrite_post_name(focus_keyword, post_name)

    # Resultaat terugschrijven naar de DataFrame
    df.at[index, column_post_name] = new_post_name

    # Alt-tekst toevoegen aan afbeeldingen
    df = add_alt_text_to_images(df)

    # EAN naar Global Unique ID kopiëren
    df = copy_ean_to_global_unique_id(df)

    # Pauze tussen API-aanroepen om rate limits te respecteren
    time.sleep(1)

# Schrijf de resultaten naar een nieuw Excel-bestand
output_file = 'herschreven_excel/bkn-living/updated_bknliving.xlsx'
df.to_excel(output_file, index=False)

print("Verwerking voltooid! Resultaten zijn opgeslagen in:", output_file)
