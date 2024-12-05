import os
from dotenv import load_dotenv
import pandas as pd
from openai import OpenAI
import re  
import time

# OpenAI API key laden
load_dotenv()
openai_api_key = os.getenv('OPENAI_API_KEY')
client = OpenAI(api_key=openai_api_key)

# Excel bestand inlezen
df = pd.read_excel('excel/autoinkoop/auto10.xlsx')

column_content = 'Content'
column_meta_title = '_yoast_wpseo_title'
column_meta_description = '_yoast_wpseo_metadesc'
column_focus_keyword = '_yoast_wpseo_focuskw'
column_title = 'Title'

max_title_length = 60  
max_description_length = 150  
company_name = "AutoInkoopService"

# Functie om focus keyword te genereren indien de kolom leeg is
def generate_focus_keyword(title):
    if pd.notnull(title):
        # Pak de eerste twee woorden van de title, indien mogelijk
        words = title.split()
        focus_keyword = ' '.join(words[:2])  # Gebruik de eerste twee woorden als focus keyword
        return focus_keyword
    return "algemeen"

# Functie om WordPress-shortcodes om te zetten naar HTML, met speciale behandeling voor images
def convert_shortcodes_to_html(content):
    # Zet de image shortcodes om naar een <img> tag
    content = re.sub(r'\[vcex_image[^\]]*image_id="(\d+)"[^\]]*\]', r'<img src="path/to/images/\1.jpg" alt="Auto Inkoop Service" />', content)

    # Verwijder alle andere shortcodes (maar behoud de afbeeldingen)
    content = re.sub(r'\[/?(?!vcex_image)[a-zA-Z0-9_]+.*?\]', '', content)  # Verwijder andere shortcodes, maar laat afbeeldingen intact
    return content

# Functie om tekst te splitsen op basis van alinea's
def split_text_by_paragraphs(text, max_paragraphs):
    paragraphs = text.split("\n\n")
    return ['\n\n'.join(paragraphs[i:i + max_paragraphs]) for i in range(0, len(paragraphs), max_paragraphs)]

# Functie om content te herschrijven met dynamisch splitsen en contextbehoud
def rewrite_content(content, focus_keyword):
    if pd.notnull(content):
        try:
            # Eerst de shortcodes omzetten naar HTML
            content = convert_shortcodes_to_html(content)
            
            word_count = len(content.split())
            if word_count > 1000:
                split_contents = split_text_by_paragraphs(content, max_paragraphs=10)
            else:
                split_contents = [content]

            rewritten_parts = []
            for idx, part in enumerate(split_contents):
                context_text = "\n\n".join(rewritten_parts[-1:]) if rewritten_parts else ""
                combined_content = f"{context_text}\n\n{part}"

                messages = [
                    {"role": "user", "content": (
                        "Herschrijf de volgende webpagina-inhoud voor de nieuwe werkwijze van AutoInkoopService.nl, Inclusief nieuwe app-stappen en verbeterde klantenservice. "
                        "Behoud de structuur, gebruik alleen noodzakelijke tags zoals `<h2>`, `<h3>`, en `<p>` voor de structuur,"
                        "Houd het eenvoudig en overzichtelijk. Vermijd overmatige styling en herhaling van CTA's; zorg er alleen voor dat de kernstappen en belangrijke informatie goed georganiseerd zijn in een logische structuur. "
                        "Maak gebruik van <h2> en <h3> voor belangrijke secties, en <p> voor paragrafen."
                        f"\n\n{combined_content}"
                    )}
                ]

                response = client.chat.completions.create(
                    model="gpt-3.5-turbo",
                    messages=messages,
                    max_tokens=4000,
                    temperature=0.7
                )
                rewritten_text = response.choices[0].message.content.strip()
                rewritten_parts.append(rewritten_text)

            rewritten_content = ' '.join(rewritten_parts)
            return add_focus_keyword(rewritten_content, focus_keyword)

        except Exception as e:
            print(f"Error rewriting content: {e}")
            return content
    return content

# Functie om focus keyword in content te injecteren
def add_focus_keyword(content, focus_keyword):
    # Voeg het focus keyword toe op verschillende plekken in de tekst als het nog niet aanwezig is
    if content.count(focus_keyword) < 3:
        content += f"\n\n{focus_keyword}" * (3 - content.count(focus_keyword))
    return content

# Functie om content aan te vullen tot minimaal 500 woorden
def fill_content(content):
    if pd.notnull(content):
        try:
            while len(content.split()) < 500:
                prompt = (
                    "Maak de volgende tekst compleet tot 500 woorden, passend bij de inhoud: "
                    f"\n\n{content}"
                )
                response = client.chat.completions.create(
                    model="gpt-3.5-turbo",
                    messages=[{"role": "user", "content": prompt}],
                    max_tokens=2000,
                    temperature=0.7
                )
                additional_text = response.choices[0].message.content.strip()
                content += " " + additional_text
            return content

        except Exception as e:
            print(f"Error filling content: {e}")
            return content
    return content

def generate_meta_title(subject, focus_keyword):
    prompt = (
        f"Schrijf een SEO-geoptimaliseerde meta title voor '{subject}', beginnend met het focus keyword '{focus_keyword}'. "
        "Gebruik maximaal 5-6 woorden en voeg geen herhalingen van het focus keyword toe. "
        "De titel mag niet langer zijn dan 55 karakters, inclusief spaties. "
        f"Bijvoorbeeld: '{focus_keyword} - Uw auto verkopen in één stap'."
        f"Ander voorbeeld: '{focus_keyword} - inkoop auto | '{company_name}'."
    )
    try:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=60,
        )
        
        # Verwerk de gegenereerde titel
        meta_title = response.choices[0].message.content.strip()
        
        # Haal aanhalingstekens weg en controleer de lengte
        meta_title = meta_title.replace('"', '').replace("'", "")
        
        # Voeg bedrijfsnaam toe als de titel kort genoeg is
        if len(meta_title) < 40:
            full_title = f"{meta_title} | {company_name}"
        else:
            full_title = meta_title

        return full_title
    except Exception as e:
        print(f"Error generating meta title: {e}")
        return ""

def generate_meta_description(subject, focus_keyword, existing_description):
    if focus_keyword.lower() in subject.lower():
        subject = subject.replace(focus_keyword, "").strip()  # Vermijd dubbele focus keyword in description
    prompt = (
        f"Schrijf een SEO-geoptimaliseerde meta description voor {subject}, beginnend met '{focus_keyword}'. "
        "De description mag maximaal 150 karakters bevatten. Ik wil maximaal 2 zinnen: "
        f"1) Begin met '{focus_keyword}' en geef de belangrijkste boodschap aan. "
        "2) Voeg een heldere call-to-action toe."
    )
    try:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=100,
            temperature=0.6
        )

        description = response.choices[0].message.content.strip()
        
        # Opsplitsen in zinnen en houd maximaal twee zinnen
        sentences = description.split('. ')
        if len(sentences) > 2:
            description = '. '.join(sentences[:2]) + '.'
        
        # Limiteer tot maximaal 150 tekens
        if len(description) > 150:
            description = description[:150].rstrip()  # Limiteer tot 150 karakters en verwijder overtollige spaties
            if description[-1] != '.':  # Zorg ervoor dat de description correct wordt afgesloten
                description += '.'
        
        return description
    except Exception as e:
        print(f"Error generating meta description: {e}")
        return existing_description

# Verwerking van elke rij in de DataFrame
new_meta_titles = []
new_meta_descriptions = []

for idx, row in df.iterrows():
    print(f"Processing row {idx + 1} of {len(df)}")
    
    # Focus keyword genereren indien nodig
    focus_keyword = row[column_focus_keyword]
    if pd.isnull(focus_keyword):
        focus_keyword = generate_focus_keyword(row[column_title])
        df.at[idx, column_focus_keyword] = focus_keyword  # Sla de gegenereerde focus keyword op
    
    # Content herschrijven, aanvullen en focus keyword toevoegen
    original_content = row[column_content]
    rewritten_content = rewrite_content(original_content, focus_keyword)
    filled_content = fill_content(rewritten_content)
    df.at[idx, column_content] = filled_content

    # Meta title en description genereren
    subject = row[column_meta_title] if pd.notnull(row[column_meta_title]) else "Geen onderwerp"
    existing_description = row[column_meta_description] if pd.notnull(row[column_meta_description]) else ""

    new_meta_title = generate_meta_title(subject, focus_keyword)
    new_meta_description = generate_meta_description(subject, focus_keyword, existing_description)

    new_meta_titles.append(new_meta_title)
    new_meta_descriptions.append(new_meta_description)

# Voeg de nieuwe meta titles en descriptions toe aan de DataFrame
df[column_meta_title] = new_meta_titles
df[column_meta_description] = new_meta_descriptions

# Sla de gewijzigde DataFrame op in een nieuw Excel-bestand
output_file = 'herschreven_excel/autoinkoop/herschreven_auto10.xlsx'
df.to_excel(output_file, index=False)

print("De content, meta titles en descriptions zijn herschreven en opgeslagen.")
