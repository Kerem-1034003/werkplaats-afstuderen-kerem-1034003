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
    Kies een enkel woord dat het meest specifieke en relevante aspect van het product beschrijft.
    Het keyword moet specifiek zijn voor het product, maximaal 1 woord bevatten, en in het Nederlands zijn.
    Vermijd reeds gebruikte keywords en zorg voor variatie. Gebruik het meest specifieke en relevante woord dat het product in de titel '{post_title}' beschrijft.
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

    # Controleer of het keyword meer dan 2 woorden bevat
    words = keyword.split()
    if len(words) > 1:
        keyword = " ".join(words[:1])  # Beperk tot de eerste 2 woorden
        
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

# Functie voor de productbeschrijving
def rewrite_product_content(post_content, focus_keyword, new_title):
    try:
        prompt = f"""
        Schrijf een uitgebreide HTML-geformatteerde productbeschrijving van minimaal 300 woorden over het product '{new_title}', 
        met headings en paragrafen.
        Gebruik '{focus_keyword}' maximaal 7 keer op een natuurlijk manier in de tekst. Als de limiet van 7 is bereikt, gebruik dan alternatieve formuleringen. 
        Beschrijf de functies, voordelen, en specificaties op een klantgerichte manier.
        Verbeter de volgende productbeschrijving zodat deze voldoet aan de volgende criteria:
        - Voeg ten minste één geordende of ongeordende lijst toe.
        - Gebruik signaalwoorden (zoals 'daarom', 'hierdoor', 'bovendien', 'echter', 'ten slotte') in ten minste 30% van de zinnen.
        - Zorg ervoor dat de eerste heading een <h3>-tag is en de subkoppen een <h4>-tag.
        Ik verwacht een product beschrijving terug van minimaal 300 woorden. en nier onder 300 woorden. 
        
        Gebruik de volgende beschrijving als uitgangspunt: '{post_content}'.
        """

        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=1500
        )
        
        new_content = response.choices[0].message.content.strip()

         # Controleer of er al een <h3>-heading bestaat
        if "<h3>" in new_content:
            # Vervang de eerste <h3>-heading met de product title
            new_content = re.sub(r'<h3>(.*?)<\/h3>', f"<h3>{new_title}</h3>", new_content, count=1)

        # Verwijder ongewenste heading-symbolen zoals ### (indien van toepassing)
        new_content = re.sub(r'#+\s?', '', new_content)
        # Vervang de eerste heading door <h3> en de subkoppen door <h4>
        new_content = re.sub(r'(<h1>)', '<h3>', new_content)  # Eerste kop naar <h3>
        new_content = re.sub(r'(<h2>)', '<h4>', new_content)  # Subkoppen naar <h4>

        # Controleer en corrigeer gebruik van het focus keyword
        focus_count = new_content.lower().count(focus_keyword.lower())
        if focus_count > 6:
            print(f"Focus keyword '{focus_keyword}' komt {focus_count} keer voor. Content wordt aangepast.")

            adjustment_prompt = f"""
            De volgende tekst bevat de focus keyword '{focus_keyword}' {focus_count} keer, wat te veel is. 
            Herwerk de tekst zodat het keyword maximaal 6 keer voorkomt. Hier is de originele tekst: '{new_content}'.
            """

            adjustment_response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": adjustment_prompt}],
                max_tokens=1000
            )
            
            adjusted_content = adjustment_response.choices[0].message.content.strip()
            new_content = adjusted_content

             # Controleer of er al een <h3>-heading bestaat
        if "<h3>" in new_content:
            # Vervang de eerste <h3>-heading met de product title
            new_content = re.sub(r'<h3>(.*?)<\/h3>', f"<h3>{new_title}</h3>", new_content, count=1)

        # Controleer of de uiteindelijke tekst aan de lengtevereisten voldoet
        word_count = len(new_content.split())
        if word_count < 200:
            print(f"Tekst bevat slechts {word_count} woorden. Extra informatie wordt toegevoegd.")
            
            additional_prompt = f"""
            Voeg meer informatie toe aan de volgende beschrijving om deze minstens 200 woorden te maken en maximaal 300 woorden. 
            Zorg ervoor dat het focus keyword '{focus_keyword}' maximaal 6 keer voorkomt, en gebruik andere relevante keywords waar nodig: '{new_content}'.
            """
            
            additional_response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": additional_prompt}],
                max_tokens=1000
            )
            
            additional_content = additional_response.choices[0].message.content.strip()
            new_content += " " + additional_content

             # Controleer of er al een <h3>-heading bestaat
        if "<h3>" in new_content:
            # Vervang de eerste <h3>-heading met de product title
            new_content = re.sub(r'<h3>(.*?)<\/h3>', f"<h3>{new_title}</h3>", new_content, count=1)

        # Retourneer de uiteindelijke herschreven beschrijving
        return new_content

    except Exception as e:
        print(f"Error rewriting product content: {e}")
        return post_content

# Itereer door elke rij in de DataFrame
for index, row in df.iterrows():
    print(f"Processing row {index + 1} of {len(df)}") 
    
    post_title = row[column_post_title]
    # Stap 1: Vertaal de titel naar Nederlands
    post_title = translate_title_to_dutch(post_title)
    df.at[index, column_post_title] = post_title  # Update de titel in de DataFrame

    # Stap 2: Genereer één focus keyword
    focus_keyword = generate_focus_keyword(post_title)
    df.at[index, column_focus_keyword] = focus_keyword

    # Stap 3: Producttitel herschrijven
    new_title = rewrite_product_title(post_title, focus_keyword)
    df.at[index, column_post_title] = new_title

    # Stap 5: Productbeschrijving herschrijven
    post_content = row[column_post_content]
    new_content = rewrite_product_content(post_content, focus_keyword, new_title)
    df.at[index, column_post_content] = new_content

    # Pauze tussen API-aanroepen om rate limits te respecteren
    time.sleep(0.5)
   
# Schrijf de resultaten naar een nieuw Excel-bestand
output_file = 'herschreven_excel/simpledeal/simple.xlsx'
df.to_excel(output_file, index=False)

print("Verwerking voltooid! Resultaten zijn opgeslagen in:", output_file)