import os
from dotenv import load_dotenv
import pandas as pd
import openai

# OpenAI API key laden
load_dotenv()
openai_api_key = os.getenv('OPENAI_API_KEY')
openai.api_key = openai_api_key

df = pd.read_excel('excel/bkn-living/bknliving10.xlsx')

# Definieer de kolomnamen
column_post_title = 'post_title'
column_post_content = 'post_content'
column_meta_title = 'meta:rank_math_title'
column_meta_description = 'meta:rank_math_description'
column_focus_keyword = 'meta:rank_math_focus_keyword'
column_alg_ean = 'meta:_alg_ean'
column_global_unique_id = 'meta:_global_unique_id'
column_images = 'images'

company_name = "Bkn Living"

def improve_or_generate_focus_keyword(post_title, current_focus_keyword=None):
    try:
        # Bepaal het prompt op basis van of er al een focus keyword is
        if isinstance(current_focus_keyword, str) and len(current_focus_keyword.strip()) > 0:
            prompt = f"""
            Verbeter het volgende focus keyword en houd het binnen 60 karakters, inclusief spaties:'{current_focus_keyword}'.
            Zorg dat het keyword is gebaseerd op het producttype, merk, en unieke specificaties zoals kleur of materiaal.
            Het product heeft de zoekwoord: '{post_title}'.
            Dit keyword zal worden gebruikt voor SEO, dus zorg dat het professioneel is en geschikt voor titels en beschrijvingen.
            Zorg ervoor dat de keyword in correct Nederlands is geschreven. dus géén Vlaamse woorden.
            """
        else:
            prompt = f"""
            Bepaal een nieuw SEO focus keyword voor het volgende product op basis van het producttype, merk en unieke specificaties zoals kleur of materiaal.
            Zorg dat het keyword niet langer is dan 60 karakters, inclusief spaties.
            haal het focus keyword idee uit de kolom: '{post_title}'.
            Zorg ervoor dat het keyword professioneel is en geschikt voor titels en beschrijvingen.
            Zorg ervoor dat de keyword in correct Nederlands is geschreven. dus géén Vlaamse woorden.
            """
        
        # Vraag OpenAI om een keyword te genereren of verbeteren
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "user", "content": prompt}
            ],
            max_tokens=30  # Houd de response kort en bondig
        )
        
        # Verwerk de gegenereerde focus keyword
        focus_keyword = response['choices'][0]['message']['content'].strip()

        # Verwijder ongewenste aanhalingstekens
        focus_keyword = focus_keyword.replace('"', '').replace("'", "")

        return focus_keyword

    except Exception as e:
        print(f"Error generating focus keyword: {e}")
        return current_focus_keyword or post_title  # Als er iets misgaat, geef de huidige waarde terug
    
def rewrite_product_title(post_title, focus_keyword):
    try:
        prompt = f"""
        Schrijf een producttitel van maximaal 60 tekens, gebruik het focus keyword '{focus_keyword}' als idee, gevolgd door een power word en de belangrijkste specificatie (kleur of materiaal).
        Het originele product heeft de naam: '{post_title}'.
        Zorg ervoor dat de titel in correct Nederlands is geschreven. dus géén Vlaamse woorden.
        """
        
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=60  # Houd de response kort en bondig
        )
        
        new_title = response['choices'][0]['message']['content'].strip()
        return new_title.replace('"', '').replace("'", "")  # Verwijder ongewenste aanhalingstekens

    except Exception as e:
        print(f"Error rewriting product title: {e}")
        return post_title  # Geef het originele titel terug bij een fout

def rewrite_product_content(post_content, focus_keyword):
    try:
        prompt = f"""
        Schrijf een productbeschrijving van minimaal 250 woorden die het focus keyword '{focus_keyword}' 2-3 keer op een natuurlijke manier opneemt. 
        Beschrijf de belangrijkste functies, voordelen en unieke kenmerken van het product in een menselijke toon en voeg indien relevant specificaties toe.
        Het originele product heeft de beschrijving: '{post_content}'.
        Zorg ervoor dat de product beschrijving in correct Nederlands is geschreven. dus géén Vlaamse woorden.
        """
        
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=600  # Houd de response kort en bondig
        )
        
        new_content = response['choices'][0]['message']['content'].strip()
        return new_content.replace('"', '').replace("'", "")  # Verwijder ongewenste aanhalingstekens

    except Exception as e:
        print(f"Error rewriting product content: {e}")
        return post_content  # Geef de originele beschrijving terug bij een fout

def generate_meta_title(post_title):
    prompt = (
        f"Schrijf een korte en krachtige SEO-geoptimaliseerde meta title."
        f"Het zoekwoord is '{post_title}'."
        "De titel moet bestaan uit 5-6 woorden, en mag niet langer zijn dan 50 karakters inclusief spaties."
        "Zorg ervoor dat de titel niet wordt afgebroken en aantrekkelijk is voor de lezer."
        "Bijvoorbeeld: 'Luxe draadtafel | Bkn living'."
        "verwijde dit soort tekens '%%title%%', '%%sitename%%', '%%page%%','%%sep%%'."
        "Zorg ervoor dat de meta title in correct Nederlands is geschreven. dus géén Vlaamse woorden."
    )

    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}],
        max_tokens = 30
    )

    meta_title = response['choices'][0]['message']['content'].strip()

    meta_title = meta_title.replace('"','' ).replace("'","")
    
    if len(meta_title) < 47:
        full_title = f"{meta_title} | {company_name}"
    else:
        full_title = meta_title

    return full_title

def generate_meta_description(post_title, focus_keyword, current_meta_description=None):
    try:
        # Bepaal het prompt op basis van of er al een meta description is
        if isinstance(current_meta_description, str) and len(current_meta_description.strip()) > 0:
            prompt = f"""
            Verbeter de volgende meta description en houd het binnen 150 karakters, inclusief spaties:
            '{current_meta_description}'.
            Zorg voor een unieke en originele openingszin die de lezer nieuwsgierig maakt. Vermijd het gebruik van standaardopeningen zoals "Ontdek" en "Verken".
            De description moet de belangrijkste functies, voordelen en unieke kenmerken van het product behouden.
            Het focus keyword is '{focus_keyword}' en het product heeft de titel '{post_title}'.
            Zorg ervoor dat de meta beschrijving in correct Nederlands is geschreven.
            """
        else:
            prompt = f"""
            Schrijf een aantrekkelijke en SEO-geoptimaliseerde meta description van maximaal 150 karakters. 
            Het focus keyword is '{focus_keyword}' en het product heeft de titel '{post_title}'.
            Creëer een originele openingszin die de aandacht trekt en vermijd standaardopeningen zoals "Ontdek" en "Verken".
            Beschrijf de belangrijkste functies en voordelen van het product in een menselijke toon.
            Ik wil maximaal 2 zinnen in de description:
            - De eerste zin moet de belangrijkste boodschap van het product duidelijk maken.
            - De tweede zin moet een duidelijke call-to-action bevatten.
            Zorg ervoor dat de meta beschrijving in correct Nederlands is geschreven.
            """
        
        # Vraag OpenAI om een meta description te genereren of verbeteren
        response = openai.ChatCompletion.create(
            model='gpt-3.5-turbo',
            messages=[{"role": "user", "content": prompt}],
            max_tokens=150  # Houd het kort en bondig
        )

        # Ontvang en verwerk de gegenereerde meta description
        meta_description = response['choices'][0]['message']['content'].strip()

        # Verwijder ongewenste aanhalingstekens
        meta_description = meta_description.replace('"', '').replace("'", "")

        # Opsplitsen in zinnen en houd maximaal twee zinnen
        sentences = meta_description.split('. ')
        if len(sentences) > 2:
            meta_description = '. '.join(sentences[:2]) + '.'

        # Check of de lengte meer dan 150 karakters is
        if len(meta_description) > 150:
            # Knip de tekst af na de laatste volledige zin binnen de limiet
            shortened_description = ''
            for sentence in sentences:
                if len(shortened_description) + len(sentence) + 1 <= 150:  # +1 voor de punt
                    shortened_description += sentence + '. '
                else:
                    break
            # Verwijder de laatste spatie en punt als de lengte te kort is
            meta_description = shortened_description.strip()

        return meta_description

    except Exception as e:
        print(f"Error generating meta description: {e}")
        # Geef een standaard waarde terug als er een fout optreedt
        return f"Product beschrijving van {post_title} met focus op {focus_keyword}."

# Nieuwe functie om ean naar global_unique_id te kopiëren
def copy_ean_to_global_unique_id(df, ean_column='meta:_alg_ean', global_unique_id_column='meta:_global_unique_id'):
    """
    Kopieert de waarde van de kolom 'meta:_alg_ean' naar de kolom 'meta:_global_unique_id'
    als er een waarde aanwezig is in 'meta:_alg_ean'.
    """
    for idx, row in df.iterrows():
        ean_value = row[ean_column]
        if pd.notna(ean_value):  # Controleer of er een waarde aanwezig is in de kolom 'meta:_alg_ean'
            df.at[idx, global_unique_id_column] = ean_value  # Kopieer de waarde naar 'meta:_global_unique_id'
    return df

# Functie om alt-tekst aan de eerste afbeelding in de kolom 'images' toe te voegen
def add_alt_text_to_images(df, images_column='images', title_column='post_title', focus_keyword_column='meta:rank_math_focus_keyword', meta_title_column='meta:rank_math_title'):
    """
    Voeg een AI-gegenereerde alt-tekst toe aan alleen de eerste afbeelding in de lijst van afbeeldingen in de kolom 'images'.
    De AI gebruikt 'post_title', 'focus_keyword' en 'meta_title' als inspiratie voor de alt-tekst.
    """
    for idx, row in df.iterrows():
        if pd.notna(row[images_column]):
            images = row[images_column].split(' | ')  # Splits de afbeeldingen in de lijst
            if images:
                # Stel de prompt samen met beschikbare kolomwaarden
                prompt = f"""
                Genereer een korte en beschrijvende alt-tekst van ongeveer 30 karakters voor een productafbeelding.
                Gebruik één van deze referenties: post_title: '{row[title_column]}', 
                focus_keyword: '{row[focus_keyword_column]}', of meta_title: '{row[meta_title_column]}'.
                Houd de beschrijving relevant en helder, en in correct Nederlands.
                """
                
                try:
                    # Vraag OpenAI om een alt-tekst te genereren
                    response = openai.ChatCompletion.create(
                        model="gpt-3.5-turbo",
                        messages=[{"role": "user", "content": prompt}],
                        max_tokens=15  # Houd de response kort en bondig
                    )
                    
                    # Ontvang en verwerk de gegenereerde alt-tekst
                    alt_text = response['choices'][0]['message']['content'].strip()
                    
                    # Voeg de alt-tekst toe aan alleen de eerste afbeelding
                    images[0] += f" ! alt : {alt_text}"
                
                except Exception as e:
                    print(f"Error generating alt text for row {idx}: {e}")
                
                # Zet de lijst met afbeeldingen weer samen in de oorspronkelijke kolomindeling
                df.at[idx, images_column] = ' | '.join(images)
                
    return df

# Loop door de DataFrame en verbeter of genereer
for idx, row in df.iterrows():
    # Focus keyword
    current_focus_keyword = row[column_focus_keyword] if pd.notna(row[column_focus_keyword]) else None
    new_focus_keyword = improve_or_generate_focus_keyword(
        post_title=row[column_post_title],
        current_focus_keyword=current_focus_keyword
    )
    # Update de focus keyword in de bestaande kolom
    df.at[idx, column_focus_keyword] = new_focus_keyword

    # Herschrijf de producttitel
    new_title = rewrite_product_title(row[column_post_title], new_focus_keyword)
    df.at[idx, column_post_title] = new_title

    # Herschrijf de productbeschrijving
    new_content = rewrite_product_content(row[column_post_content], new_focus_keyword)
    df.at[idx, column_post_content] = new_content

    post_title = row[column_post_title]
    # Genereer en update de meta title
    new_meta_title = generate_meta_title(post_title)
    df.at[idx, column_meta_title] = new_meta_title
    
    current_meta_description = row[column_meta_description] if pd.notna(row[column_meta_description]) else None
    new_meta_description = generate_meta_description(
        post_title=row[column_post_title], 
        focus_keyword=row[column_focus_keyword], 
        current_meta_description=current_meta_description
    )
    df.at[idx, column_meta_description] = new_meta_description

# Kopieer de EAN naar Global Unique ID
df = copy_ean_to_global_unique_id(df)

# Voeg de alt-tekst toe aan afbeeldingen
df = add_alt_text_to_images(df)

# Opslaan in hetzelfde bestand of een nieuw bestand als je dat wilt controleren
df.to_excel('herschreven_excel/bkn-living/updated_bkn-living.xlsx', index=False)
