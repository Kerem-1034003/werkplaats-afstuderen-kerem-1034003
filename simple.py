import os
import re
from dotenv import load_dotenv
import pandas as pd
from openai import OpenAI
import time

load_dotenv()

openai_api_key = os.getenv('OPENAI_API_KEY')

# Initialiseer de OpenAI client
client = OpenAI(api_key=openai_api_key)

# Laad de DataFrame
df = pd.read_excel('excel/simpledeal/Map2.xlsx')

# Definieer de kolomnamen
column_post_title = 'post_title'
column_post_name = 'post_name'
column_post_content = 'post_content'
column_meta_title = 'meta:_yoast_wpseo_title'
column_meta_description = 'meta:_yoast_wpseo_metadesc'
column_focus_keyword = 'meta:_yoast_wpseo_focuskw'
column_images = 'images'

# Cast relevante kolommen naar strings om datatypeproblemen te voorkomen
df[column_focus_keyword] = df[column_focus_keyword].astype(str)
df[column_meta_title] = df[column_meta_title].astype(str)
df[column_meta_description] = df[column_meta_description].astype(str)
df[column_post_name] = df[column_post_name].astype(str)

company_name = "Simpledeal"

# Functie om Duitse titels naar het Nederlands te vertalen
def translate_title_to_dutch(title):
    try:
        prompt = f"""
        Vertaal de volgende producttitel van het Duits naar correct Nederlands: '{title}'. 
        Zorg ervoor dat de vertaling duidelijk en klantgericht is.
        """
        
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=100
        )
        
        translated_title = response.choices[0].message.content.strip()
        return translated_title
    except Exception as e:
        print(f"Error translating title: {e}")
        return title  # Retourneer de originele titel als de vertaling faalt

# Functie voor het genereren van één focus keyword
def generate_focus_keyword(post_title):
    prompt = f"""
    Genereer één SEO-geoptimaliseerd focus keyword voor '{post_title}'. 
    Vermijd reeds gebruikte keywords en zorg voor variatie. Gebruik het meest specifieke en relevante woord dat het product in de titel '{post_title}' beschrijft.
    Het keyword moet specifiek zijn voor het product, maximaal 2 woorden bevatten, en in het Nederlands zijn.
    Ik wil Geen littekens, voorzetsels of kleuren in de focus keyword zien.
    Vermijd speciale tekens of afkortingen.
    """

    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=10
    )

    keyword = response.choices[0].message.content.strip().replace('"', '').replace("'", "")
    keyword = re.sub(r'[^a-zA-Z0-9\s-]', '', keyword).lower()  # Verwijder speciale tekens en zet alles in kleine letters
    return keyword

# Functie voor het herschrijven van de producttitel
def rewrite_product_title(post_title, focus_keyword):
    try:
        prompt = f"""
        Schrijf een producttitel beginnend met het focus keyword '{focus_keyword}', en gevolgd door een powerword
        De titel moet bestaan uit 6-7 woorden, en mag niet langer zijn dan 80 karakters inclusief spaties. 
        Het originele product heeft de naam: '{post_title}'.
        Zorg ervoor dat de titel in correct Nederlands is geschreven. dus géén Duitse woorden.
        """
        
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=60
        )
        
        new_title = response.choices[0].message.content.strip().replace('"', '').replace("'", "")
        return new_title
    except Exception as e:
        print(f"Error rewriting product title: {e}")
        return post_title

# Functie voor het herschrijven van de URL (post_name)
def rewrite_post_name(focus_keyword, post_name, max_retries=3):
    prompt = f"""
    Herschrijf de URL '{post_name}' zodat deze begint met het focus keyword '{focus_keyword}', 
    bevat 5-6 woorden en niet meer dan 70 karakters inclusief spaties. Gebruik alleen letters, cijfers, en koppeltekens.
    Geen speciale tekens of slashes (/). 
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

# Functie voor het herschrijven van de productbeschrijving
def rewrite_product_content(post_content, focus_keyword):
    try:
        # Herschrijven van de productbeschrijving volgens de SEO-vereisten
        prompt = f"""
        Schrijf een HTML-geformatteerde productbeschrijving van minimaal 350 woorden voor het product, 
        met alleen headings en korte paragrafen. Gebruik de focus keywords '{focus_keyword}' op een natuurlijke manier. 
        Gebruik '{focus_keyword}' maximaal 5 keer in de tekst. Als de limiet van 5 is bereikt, gebruik dan alternatieve formuleringen. 
        Beschrijf de functies, voordelen, en specificaties op een klantgerichte manier.
        Verbeter de volgende productbeschrijving zodat deze voldoet aan de volgende criteria:
        - Voeg ten minste één geordende of ongeordende lijst toe.
        - Zorg dat geen enkele sectie langer is dan 300 woorden zonder subkopteksten.
        - Beperk zinnen tot maximaal 20 woorden en verbeter de leesbaarheid.
        - Gebruik signaalwoorden (zoals 'daarom', 'hierdoor', 'bovendien', 'echter', 'ten slotte') in ten minste 30% van de zinnen.
        - Gebruik het focus keyword '{focus_keyword}' maximaal 5 keer in de tekst.
        - Zorg ervoor dat de eerste heading een <h3>-tag is en de subkoppen een <h4>-tag.
        
        Gebruik de volgende beschrijving als uitgangspunt: '{post_content}'.
        """

        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=1500
        )
        
        new_content = response.choices[0].message.content.strip()

        # Verwijder ongewenste heading-symbolen zoals ### (indien van toepassing)
        new_content = re.sub(r'#+\s?', '', new_content)

        # Vervang de eerste heading door <h3> en de subkoppen door <h4>
        new_content = re.sub(r'(<h1>)', '<h3>', new_content)  # Eerste kop naar <h3>
        new_content = re.sub(r'(<h2>)', '<h4>', new_content)  # Subkoppen naar <h4>

        # Controleer en corrigeer gebruik van het focus keyword
        focus_count = new_content.lower().count(focus_keyword.lower())
        if focus_count > 5:
            print(f"Focus keyword '{focus_keyword}' komt {focus_count} keer voor. Content wordt aangepast.")

            adjustment_prompt = f"""
            De volgende tekst bevat de focus keyword '{focus_keyword}' {focus_count} keer, wat te veel is. 
            Herwerk de tekst zodat het keyword maximaal 5 keer voorkomt. Hier is de originele tekst: '{new_content}'.
            """

            adjustment_response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": adjustment_prompt}],
                max_tokens=1000
            )
            
            adjusted_content = adjustment_response.choices[0].message.content.strip()
            new_content = adjusted_content

        # Controleer of de uiteindelijke tekst aan de lengtevereisten voldoet
        word_count = len(new_content.split())
        if word_count < 300:
            print(f"Tekst bevat slechts {word_count} woorden. Extra informatie wordt toegevoegd.")
            
            additional_prompt = f"""
            Voeg meer informatie toe aan de volgende beschrijving om deze minstens 300 woorden te maken en maximaal 500 woorden. 
            Zorg ervoor dat het focus keyword '{focus_keyword}' maximaal 5 keer voorkomt, en gebruik andere relevante keywords waar nodig: '{new_content}'.
            """
            
            additional_response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": additional_prompt}],
                max_tokens=1000
            )
            
            additional_content = additional_response.choices[0].message.content.strip()
            new_content += " " + additional_content

        # Retourneer de uiteindelijke herschreven beschrijving
        return new_content

    except Exception as e:
        print(f"Error rewriting product content: {e}")
        return post_content

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
        Schrijf een SEO-geoptimaliseerde meta description van maximaal 150 tekens beginnend met het focus'{focus_keyword}'. 
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

# Set voor unieke gegenereerde keywords voor de hele DataFrame
generated_keywords = set()

# Itereer door elke rij in de DataFrame
for index, row in df.iterrows():
    print(f"Processing row {index + 1} of {len(df)}") 
    post_title = row[column_post_title]

    # Stap 1: Vertaal de titel naar Nederlands
    post_title = translate_title_to_dutch(post_title)
    df.at[index, column_post_title] = post_title  # Update de titel in de DataFrame

    # Stap 2: Genereer één focus keyword
    primary_keyword = generate_focus_keyword(post_title)
    df.at[index, column_focus_keyword] = primary_keyword

    # Stap 3: Producttitel herschrijven
    new_title = rewrite_product_title(post_title, primary_keyword)
    df.at[index, column_post_title] = new_title

    # Stap 4: URL herschrijven
    post_name = row[column_post_name]
    new_post_name = rewrite_post_name(primary_keyword, post_name)
    df.at[index, column_post_name] = new_post_name

    # Stap 5: Productbeschrijving herschrijven
    post_content = row[column_post_content]
    new_content = rewrite_product_content(post_content, primary_keyword)
    df.at[index, column_post_content] = new_content

    # Stap 6: Meta title en description genereren
    new_meta_title = generate_meta_title(primary_keyword)
    new_meta_description = generate_meta_description(post_title, primary_keyword)
    df.at[index, column_meta_title] = new_meta_title
    df.at[index, column_meta_description] = new_meta_description

    # Pauze tussen API-aanroepen om rate limits te respecteren
    time.sleep(1)

    # Stap 7: Alt-tekst genereren voor afbeeldingen
    df = add_alt_text_to_images(
        df, 
        images_column='images', 
        focus_keyword_column=column_focus_keyword
    )
# Schrijf de resultaten naar een nieuw Excel-bestand
output_file = 'herschreven_excel/simpledeal/Map2.xlsx'
df.to_excel(output_file, index=False)

print("Verwerking voltooid! Resultaten zijn opgeslagen in:", output_file)