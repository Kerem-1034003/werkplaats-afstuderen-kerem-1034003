import os
import re
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

# Functie voor het genereren van focus keywords met check op lege rijen
def generate_focus_keywords(post_title, num_keywords=3):
    focus_keywords = []
    attempts = 0

    while len(focus_keywords) < num_keywords and attempts < num_keywords * 2:
        attempts += 1
        prompt = f"""
        Genereer een SEO-geoptimaliseerd focus keyword voor '{post_title}'.
        Het keyword moet specifiek zijn voor het product in maximaal 2 woorden. 
        het keyword moet in het Nederlands zijn. dus geen vlaamse of engels woorden. 
        Vermijd speciale tekens, bijvoorbeeld: 'tuinstoel', 'bureaustoel', enz.
        """

        # Vraag een nieuw keyword aan de AI
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=20
        )

        # Schoon het gegenereerde keyword op en verwijder speciale tekens
        keyword = response.choices[0].message.content.strip().replace('"', '').replace("'", "")
        keyword = re.sub(r'[^a-zA-Z0-9\s-]', '', keyword)  # Verwijder speciale tekens
        keyword = keyword.lower()  # Zet om naar kleine letters voor consistentie

        # Voeg het keyword toe als het niet leeg is
        if keyword:
            focus_keywords.append(keyword)
        else:
            print(f"Leeg keyword gedetecteerd voor '{post_title}', opnieuw proberen...")

    # Controleer op lege focus keywords; herhaal als er geen keywords zijn
    if not focus_keywords:
        print(f"Geen keywords gegenereerd voor '{post_title}', opnieuw proberen...")
        return generate_focus_keywords(post_title, num_keywords)

    return ", ".join(focus_keywords[:num_keywords])

# Functie voor het herschrijven van de producttitel
def rewrite_product_title(post_title, primary_keyword):
    try:
        prompt = f"""
        Schrijf een producttitel van maximaal 60 karakters inclusief spaties, beginnend met het focus keyword '{primary_keyword}',
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
def rewrite_product_content(post_content, focus_keywords, primary_keyword):
    try:
        prompt = f"""
        Schrijf een HTML-geformatteerde productbeschrijving van minimaal 300 woorden en maximaal 500 woorden voor het product, 
        met alleen headings en korte paragrafen. Gebruik de focus keywords '{focus_keywords}' op een natuurlijke manier. 
        Gebruik '{primary_keyword}' maximaal 4 keer in de tekst. Als de limiet van 4 is bereikt, gebruik dan de andere keywords of alternatieve formuleringen. 
        Gebruik de volgende beschrijving als uitgangspunt: '{post_content}'.
        Beschrijf de functies, voordelen, en specificaties op een klantgerichte manier.
        """

        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=1500
        )
        
        new_content = response.choices[0].message.content.strip().replace('"', '').replace("'", "")
        
        # Verwijder ongewenste heading-symbolen zoals ###
        new_content = re.sub(r'#+\s?', '', new_content)

        # Controleer het aantal keer dat de primary keyword voorkomt
        primary_count = new_content.lower().count(primary_keyword.lower())

        if primary_count > 4:
            print(f"Primary keyword '{primary_keyword}' komt {primary_count} keer voor. Content wordt aangepast.")
            
            # Aanpassing prompt om overmatig gebruik van de primary keyword te corrigeren
            adjustment_prompt = f"""
            De volgende tekst bevat de focus keyword '{primary_keyword}' {primary_count} keer, wat te veel is. 
            Herwerk de tekst zodat het keyword maximaal 4 keer voorkomt. Gebruik waar nodig andere focus keywords '{focus_keywords}' 
            of alternatieve formuleringen. Hier is de originele tekst: '{new_content}'.
            """

            adjustment_response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": adjustment_prompt}],
                max_tokens=1000
            )
            
            adjusted_content = adjustment_response.choices[0].message.content.strip().replace('"', '').replace("'", "")
            new_content = adjusted_content

        # Controleer of de uiteindelijke tekst aan de lengtevereisten voldoet
        word_count = len(new_content.split())
        if word_count < 300:
            additional_prompt = f"""
            Voeg meer informatie toe aan de volgende beschrijving om deze minstens 300 woorden te maken en maximaal 500 woorden. 
            Zorg ervoor dat de focus keyword '{primary_keyword}' maximaal 4 keer voorkomt, en gebruik andere relevante keywords waar nodig: '{new_content}'.
            """
            
            additional_response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": additional_prompt}],
                max_tokens=1000
            )
            
            additional_content = additional_response.choices[0].message.content.strip().replace('"', '').replace("'", "")
            new_content += " " + additional_content

        # Retourneer de uiteindelijke herschreven beschrijving
        return new_content

    except Exception as e:
        print(f"Error rewriting product content: {e}")
        return post_content

# Functie voor het genereren van de meta title
def generate_meta_title(primary_keyword):
    try:
        prompt = (
            f"Schrijf een SEO-geoptimaliseerde meta title, beginnend met het focus keyword '{primary_keyword}', en gevolgd door een powerword. "
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
        return f"{primary_keyword} | {company_name}"

# Functie voor het genereren van de meta description
def generate_meta_description(post_title, primary_keyword):
    try:
        prompt = f"""
        Schrijf een SEO-geoptimaliseerde meta description van maximaal 150 tekens voor '{post_title}', 
        met de focus op het keyword '{primary_keyword}'. De meta description moet aantrekkelijk zijn en
        exact 2 zinnen bevatten:
        - De eerste zin beschrijft het product, beginnend met het focus keyword '{primary_keyword}'.
        - De tweede zin eindigt met een duidelijke call-to-action.
        """

        # Vraag een meta description van de AI
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=200
        )

        meta_description = response.choices[0].message.content.strip().replace('"', '').replace("'", "")

        # Controleer of de meta description binnen de 150 tekens blijft
        if len(meta_description) > 150:
            # Verkort zonder de laatste zin af te breken
            sentences = meta_description.split('. ')
            if len(sentences) > 1:
                meta_description = '. '.join(sentences[:2]) + '.'  # Houd alleen de eerste twee zinnen

            # Knip de meta description af tot 150 tekens zonder woorden te breken
            if len(meta_description) > 150:
                meta_description = meta_description[:150].rsplit(' ', 1)[0] + "."

        return meta_description
    except Exception as e:
        print(f"Error generating meta description: {e}")
        return f"Product beschrijving van {post_title} met focus op {primary_keyword}."

# Set voor unieke gegenereerde keywords voor de hele DataFrame
generated_keywords = set()

# Itereer door elke rij in de DataFrame
for index, row in df.iterrows():
    print(f"Processing row {index + 1} of {len(df)}") 
    post_title = row[column_post_title]
    
    # Meerdere focus keywords genereren zonder variatie
    focus_keywords = generate_focus_keywords(post_title)
    primary_keyword = focus_keywords.split(", ")[0]  # Eerste keyword

    # Producttitel herschrijven
    new_title = rewrite_product_title(post_title, primary_keyword)
    
    # Productbeschrijving herschrijven
    post_content = row[column_post_content]
    new_content = rewrite_product_content(post_content, focus_keywords, primary_keyword)

    # Meta title en description genereren
    new_meta_title = generate_meta_title(primary_keyword)
    new_meta_description = generate_meta_description(post_title, primary_keyword)

    # Resultaten terugschrijven naar de DataFrame
    df.at[index, column_focus_keyword] = focus_keywords
    df.at[index, column_post_title] = new_title
    df.at[index, column_post_content] = new_content
    df.at[index, column_meta_title] = new_meta_title
    df.at[index, column_meta_description] = new_meta_description

    # Pauze tussen API-aanroepen om rate limits te respecteren
    time.sleep(1)

        # Sla tussentijdse resultaten op na elke 50 rijen om verlies van gegevens te voorkomen
    if (index + 1) % 50 == 0:
        temp_output_file = f'herschreven_excel/bkn-living/temp_output_{index + 1}.xlsx'
        df.to_excel(temp_output_file, index=False)
        print(f"Tussentijdse resultaten opgeslagen in: {temp_output_file}")

# Schrijf de resultaten naar een nieuw Excel-bestand
output_file = 'herschreven_excel/bkn-living/updated_bknliving1.4.xlsx'
df.to_excel(output_file, index=False)

print("Verwerking voltooid! Resultaten zijn opgeslagen in:", output_file)
