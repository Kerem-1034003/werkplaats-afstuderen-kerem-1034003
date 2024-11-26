import os
import re
from dotenv import load_dotenv
import pandas as pd
from openai import OpenAI
import time

load_dotenv()

openai_api_key = os.getenv('OPENAI_API_KEY2')

# Initialiseer de OpenAI client
client = OpenAI(api_key=openai_api_key)

# Laad de DataFrame
df = pd.read_excel('herschreven_excel/simpledeal/simple.xlsx')

column_post_name = 'post_name'
column_focus_keyword = 'meta:_yoast_wpseo_focuskw'
column_meta_title = 'meta:_yoast_wpseo_title'
column_meta_description = 'meta:_yoast_wpseo_metadesc'
column_images = 'images'

company_name = "Simpledeal"

df[column_meta_title] = df[column_meta_title].astype(str)
df[column_meta_description] = df[column_meta_description].astype(str)
df[column_post_name] = df[column_post_name].astype(str)

# Functie voor het herschrijven van de URL (post_name)
def rewrite_post_name(focus_keyword, post_name, max_retries=3):
    prompt = f"""
    Herschrijf de URL '{post_name}' zodat deze begint met het focus keyword '{focus_keyword}', 
    post name bevat maximaal 4-5 woorden en niet meer dan 70 karakters inclusief spaties. Gebruik alleen letters, cijfers, en koppeltekens.
    Geen speciale tekens of slashes (/). Ook geen afmetingen of cijfers. 
    Vertaal de url van het Duits naar correct Nederlands. 
    """

    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=100  # Beperk tot 50 tokens voor een compacte URL
            )
            # Verwijder ongewenste tekens en controleer lengte
            new_post_name = response.choices[0].message.content.strip().replace('"', '').replace("'", "")
            
            # Gebruik regex om alle tekens behalve alfanumerieke karakters en koppeltekens te verwijderen
            new_post_name = re.sub(r"[^a-zA-Z0-9\- ]", "", new_post_name)
            
            # Vervang spaties door koppeltekens en maak de tekst geschikt voor URL
            new_post_name = new_post_name.replace(" ", "-")

            # Controleer of de URL binnen de limiet van 70 karakters valt
            if len(new_post_name) <= 70:
                return new_post_name
            else:
                # Als de naam nog steeds te lang is, beperk tot de eerste 5 woorden
                words = new_post_name.split("-")
                new_post_name = "-".join(words[:5])

                # Controleer opnieuw de lengte
                if len(new_post_name) <= 70:
                    return new_post_name
                else:
                    print(f"Warning: Generated URL still too long after adjustment: '{new_post_name}'")

        except Exception as e:
            print(f"Error rewriting post name (attempt {attempt + 1}): {e}")
            time.sleep(2)  # Wacht 2 seconden voor een nieuwe poging

    # Als alle pogingen falen, geef het originele post_name terug
    return post_name

# Functie voor het genereren van de meta title
def generate_meta_title(focus_keyword):
    try:
        prompt = (
            f"Schrijf een SEO-geoptimaliseerde meta title, beginnend met het focus keyword '{focus_keyword}', en gevolgd door een powerword. "
            "De titel moet bestaan uit 5-6 woorden, en mag niet langer zijn dan 60 karakters inclusief spaties. "
            "Zorg ervoor dat de titel niet wordt afgebroken en aantrekkelijk is voor de lezer."
        )

        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=30
        )

        meta_title = response.choices[0].message.content.strip().replace('"', '').replace("'", "")
        
        if len(meta_title) < 47:
            full_title = f"{meta_title} | {company_name}"
        else:
            full_title = meta_title

        return full_title
    except Exception as e:
        print(f"Error generating meta title: {e}")
        return f"{focus_keyword} | {company_name}"

def generate_meta_description(post_title, focus_keyword):
    try:
        prompt = f"""
        Schrijf een SEO-geoptimaliseerde meta description beginnend met het focus keyword'{focus_keyword}'. 
        De meta descriptionmoet bestaan uit maximaal 2 zinnen en mag niet langer zijn dan 150 tekens inclusief spaties 
        De meta description moet aantrekkelijk zijn en exact 2 zinnen bevatten:
        - De eerste zin beschrijft het product, beginnend met het focus keyword '{focus_keyword}'.
        - De tweede zin eindigt met een duidelijke call-to-action.
        """

        # Vraag een meta description van de AI
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=200
        )

        meta_description = response.choices[0].message.content.strip().replace('"', '').replace("'", "")
        
        sentences = meta_description.split('. ')

        if len(sentences) > 2:
            meta_description = '. '.join(sentences[:2]) + '.'  # Houd alleen de eerste twee zinnen

        return meta_description
    except Exception as e:
        print(f"Error generating meta description: {e}")
        return f"Product beschrijving van {post_title} met focus op {focus_keyword}."
    
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

# Verwerking
for index, row in df.iterrows():
    print(f"Processing row {index + 1} of {len(df)}")
    
    post_name = row[column_post_name]
    focus_keyword = row[column_focus_keyword]
    
    new_post_name = rewrite_post_name(focus_keyword, post_name)
    df.at[index, column_post_name] = new_post_name

    new_meta_title = generate_meta_title(focus_keyword)
    df.at[index, column_meta_title] = new_meta_title

    new_meta_description = generate_meta_description(row['post_title'], focus_keyword)
    df.at[index, column_meta_description] = new_meta_description

    df = add_alt_text_to_images(
        df, 
        images_column='images', 
        focus_keyword_column=column_focus_keyword
    )

# Opslaan
output_file = 'herschreven_excel/simpledeal/updated_simple.xlsx'
df.to_excel(output_file, index=False)

print("Verwerking voltooid! Resultaten zijn opgeslagen in:", output_file)