import os
from dotenv import load_dotenv
import pandas as pd
from openai import OpenAI
import time

# Laad de OpenAI API key
load_dotenv()
openai_api_key_bkn = os.getenv('OPENAI_API_KEY')

# Initialiseer de OpenAI client
client = OpenAI(api_key=openai_api_key_bkn)

df = pd.read_excel('excel/bkn-living/bknliving.xlsx')

# Definieer de kolomnamen
column_post_title = 'post_title'
column_post_content = 'post_content'
column_meta_title = 'meta:rank_math_title'
column_meta_description = 'meta:rank_math_description'
column_focus_keyword = 'meta:rank_math_focus_keyword'

company_name = "Bkn Living"

# Functie voor het genereren of verbeteren van het focus keyword
# Functie voor het genereren of verbeteren van het focus keyword
def improve_or_generate_focus_keyword(post_title, current_focus_keyword=None):
    try:
        if current_focus_keyword:
            prompt = f"""
            Herschrijf het SEO-geoptimaliseerde focus keyword '{current_focus_keyword}' naar een nieuw focus keyword van maximaal 1 woord.
            Het nieuwe keyword moet het meest relevante en specifieke woord zijn dat het product beschrijft, afgeleid van de titel '{post_title}'.
            Focus op productcategorieën en vermijd synoniemen of algemene termen die niet precies zijn. 
            Voorbeelden van goede keywords zijn 'draadtafel', 'vitrinekast', 'tuinstoel'.
            Zorg ervoor dat het keyword in correct Nederlands is geschreven, dus géén Vlaamse woorden.
            """
        else:
            prompt = f"""
            Bepaal een nieuw SEO focus keyword van maximaal 1 woord.
            Het keyword moet het meest relevante en specifieke woord zijn dat het product beschrijft, afgeleid van de producttitel: '{post_title}'.
            Gebruik sleutelwoorden die typisch zijn voor het product en vermijd algemene of onduidelijke termen. 
            Denk aan voorbeelden: 'draadtafel', 'salontafel', 'vitrinekast'.
            Zorg ervoor dat het keyword professioneel is en geschikt voor titels en beschrijvingen, en in correct Nederlands is geschreven.
            """
        
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=30
        )
        
        # Haal de gegenereerde content op en verwijder aanhalingstekens
        focus_keyword = response.choices[0].message.content.strip().replace('"', '').replace("'", "")
        return focus_keyword
    except Exception as e:
        print(f"Error generating focus keyword: {e}")
        return current_focus_keyword or post_title


# Functie voor het herschrijven van de producttitel
def rewrite_product_title(post_title, focus_keyword):
    try:
        prompt = f"""
        Schrijf een producttitel van maximaal 60 karakters inclusief spaties, beginnend met het focus keyword '{focus_keyword}',
        gevolgd door een powerword.
        Het originele product heeft de naam: '{post_title}'.
        Zorg ervoor dat de titel in correct Nederlands is geschreven. dus géén Vlaamse woorden.
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

# Functie voor het herschrijven van de productbeschrijving
def rewrite_product_content(post_content, focus_keyword):
    try:
        prompt = f"""
        Schrijf een productbeschrijving van minimaal 250 woorden die het focus keyword '{focus_keyword}' 2-3 keer op een natuurlijke manier opneemt. 
        Beschrijf de belangrijkste functies, voordelen en unieke kenmerken van het product in een menselijke toon en voeg indien relevant specificaties toe.
        Het originele product heeft de beschrijving: '{post_content}'.
        Zorg ervoor dat de product beschrijving in correct Nederlands is geschreven. dus géén Vlaamse woorden.
        """
        
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=600
        )
        
        new_content = response.choices[0].message.content.strip().replace('"', '').replace("'", "")
        return new_content
    except Exception as e:
        print(f"Error rewriting product content: {e}")
        return post_content

# Functie voor het genereren van de meta title
def generate_meta_title(focus_keyword):
    try:
        prompt = (
        f"Schrijf een SEO-geoptimaliseerde meta title, beginnend met het focus keyword '{focus_keyword}', en gevolgd door een powerword."
        "De titel moet bestaan uit 5-6 woorden, en mag niet langer zijn dan 60 karakters inclusief spaties."
        "Zorg ervoor dat de titel niet wordt afgebroken en aantrekkelijk is voor de lezer."
        "verwijde dit soort tekens '%%title%%', '%%sitename%%', '%%page%%','%%sep%%'."
        "Zorg ervoor dat de meta title in correct Nederlands is geschreven. dus géén Vlaamse woorden."
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

# Functie voor het genereren van de meta description
def generate_meta_description(post_title, focus_keyword, current_meta_description=None):
    try:
        if current_meta_description:
            prompt = f"""
            Schrijf een aantrekkelijke SEO geoptimaliseerd meta description van maximaal 150 tekens lang inclusief spaties, beginnend met het focus keyword '{focus_keyword}'.
            Zorg voor een unieke en originele openingszin die de lezer nieuwsgierig maakt. Vermijd het gebruik van standaardopeningen zoals "Ontdek" en "Verken".
            De description moet de belangrijkste functies, voordelen en unieke kenmerken van het product behouden.
            Het focus keyword is '{focus_keyword}'.
            Zorg ervoor dat de meta beschrijving in correct Nederlands is geschreven.
            """
        else:
            prompt = f"""
            Schrijf een aantrekkelijke SEO geoptimaliseerd meta description van maximaal 150 tekens lang inclusief spaties, beginnend met het focus keyword '{focus_keyword}'.
            Het focus keyword is '{focus_keyword}'.
            Creëer een originele openingszin die de aandacht trekt en vermijd standaardopeningen zoals "Ontdek" en "Verken".
            Beschrijf de belangrijkste functies en voordelen van het product in een menselijke toon.
            Ik wil maximaal 2 zinnen in de description:
            - De eerste zin moet de belangrijkste boodschap van het product duidelijk maken beginnend met het focus keyword'{focus_keyword}'.
            - De tweede zin moet een duidelijke call-to-action bevatten.
            Zorg ervoor dat de meta beschrijving in correct Nederlands is geschreven.
            """

        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=200
        )

        meta_description = response.choices[0].message.content.strip().replace('"', '').replace("'", "")
        
        # Opsplitsen in zinnen
        sentences = meta_description.split('. ')

        # Houd maximaal twee zinnen
        if len(sentences) > 2:
            meta_description = '. '.join(sentences[:2]) + '.'

        # Limiteer tot 150 tekens
        if len(meta_description) > 150:
            sentences = meta_description.split('. ')
            shortened_description = '. '.join(sentences[:2]) + '.'
            meta_description = shortened_description[:150]

        return meta_description
    except Exception as e:
        print(f"Error generating meta description: {e}")
        return f"Product beschrijving van {post_title} met focus op {focus_keyword}."


# Itereer door elke rij in de DataFrame
for index, row in df.iterrows():
    print(f"Processing row {index + 1} of {len(df)}") 
    post_title = row[column_post_title]
    current_focus_keyword = row[column_focus_keyword]
    
    # Focus keyword genereren of verbeteren
    new_focus_keyword = improve_or_generate_focus_keyword(post_title, current_focus_keyword)
    
    # Producttitel herschrijven
    new_title = rewrite_product_title(post_title, new_focus_keyword)
    
    # Productbeschrijving herschrijven
    post_content = row[column_post_content]
    new_content = rewrite_product_content(post_content, new_focus_keyword)

    # Meta title genereren
    new_meta_title = generate_meta_title(new_focus_keyword)

    # Meta description genereren
    current_meta_description = row[column_meta_description]
    new_meta_description = generate_meta_description(post_title, new_focus_keyword, current_meta_description)

    # Resultaten terugschrijven naar de DataFrame
    df.at[index, column_focus_keyword] = new_focus_keyword
    df.at[index, column_post_title] = new_title
    df.at[index, column_post_content] = new_content
    df.at[index, column_meta_title] = new_meta_title
    df.at[index, column_meta_description] = new_meta_description


    # Pauze tussen API-aanroepen om rate limits te respecteren
    time.sleep(1)

# Schrijf de resultaten naar een nieuw Excel-bestand
output_file = 'herschreven_excel/bkn-living/bknliving1.1.xlsx'
df.to_excel(output_file, index=False)

print("Verwerking voltooid! Resultaten zijn opgeslagen in:", output_file)
