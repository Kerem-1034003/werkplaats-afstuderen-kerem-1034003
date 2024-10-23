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
            Verbeter het volgende focus keyword en houd het binnen 60 karakters, inclusief spaties:
            '{current_focus_keyword}'.
            Zorg dat het keyword is gebaseerd op het producttype, merk, en unieke specificaties zoals kleur of materiaal.
            Het product heeft de zoekwoord: '{post_title}'.
            Dit keyword zal worden gebruikt voor SEO, dus zorg dat het professioneel is en geschikt voor titels en beschrijvingen.
            """
        else:
            prompt = f"""
            Bepaal een nieuw SEO focus keyword voor het volgende product op basis van het producttype, merk en unieke specificaties zoals kleur of materiaal.
            Zorg dat het keyword niet langer is dan 60 karakters, inclusief spaties.
            haal het focus keyword idee uit de kolom: '{post_title}'.
            Zorg ervoor dat het keyword professioneel is en geschikt voor titels en beschrijvingen.
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
    
# Opslaan in hetzelfde bestand of een nieuw bestand als je dat wilt controleren
df.to_excel('herschreven_excel/bkn-living/updated_bkn-living.xlsx', index=False)
