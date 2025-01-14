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
df = pd.read_excel('../excel/autoinkoop/auto10.xlsx')

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

# Functie om WordPress-shortcodes om te zetten naar HTML
def convert_shortcodes_to_html(content):
    # Basis shortcode filtering voor bijvoorbeeld [shortcode] en [shortcode param="value"]
    content = re.sub(r'\[/?[a-zA-Z0-9_]+.*?\]', '', content)  # Verwijder alle shortcodes tussen []
    return content

# Functie om tekst te splitsen op basis van alinea's
def split_text_by_paragraphs(text, max_paragraphs):
    paragraphs = text.split("\n\n")
    return ['\n\n'.join(paragraphs[i:i + max_paragraphs]) for i in range(0, len(paragraphs), max_paragraphs)]

# Functie om content te herschrijven met behoud van alinea's en minimaal 300 woorden
def rewrite_content(content, focus_keyword):
    if pd.notnull(content):
        try:
            # Eerst de shortcodes omzetten naar HTML
            content = convert_shortcodes_to_html(content)

            # Controleer het woordenaantal
            word_count = len(content.split())
            
            # Als minder dan 300 woorden, genereer extra inhoud binnen bestaande alinea's
            if word_count < 300:
                required_words = 300 - word_count
                expansion_prompt = (
                    f"De volgende tekst bevat minder dan 300 woorden. Vul de inhoud aan met ongeveer {required_words} woorden, "
                    "door de bestaande alinea's uit te breiden. Voeg geen nieuwe secties toe, maar werk binnen de bestaande structuur.\n\n"
                    f"Focus keyword: {focus_keyword}\n\n"
                    f"Inhoud:\n{content}"
                )
                response = client.chat.completions.create(
                    model="gpt-3.5-turbo",
                    messages=[{"role": "user", "content": expansion_prompt}],
                    max_tokens=4000,
                    temperature=0.7,
                )
                content = response.choices[0].message.content.strip()
                word_count = len(content.split())  # Update woordenaantal na uitbreiding

            # Splits de inhoud indien nodig
            if word_count > 1000:
                split_contents = split_text_by_paragraphs(content, max_paragraphs=10)
            else:
                split_contents = [content]

            # Herschrijf de inhoud in delen zonder woorden per alinea te verkorten
            rewritten_parts = []
            for idx, part in enumerate(split_contents):
                context_text = "\n\n".join(rewritten_parts[-1:]) if rewritten_parts else ""
                combined_content = f"{context_text}\n\n{part}"

                messages = [
                    {"role": "user", "content": (
                        "Herschrijf de volgende webpagina-inhoud voor de nieuwe werkwijze van new.autoInkoopService.nl, Inclusief nieuwe app-stappen en verbeterde klantenservice. Ik verwacht een breed content per alinea. Dus behoud de woordenaantal per pagina. Ik verwacht niet dat je de woordenaantal verminderd."
                        "De nieuwe werkwijze houd in dat we alle klanten naar ons app willen sturen door het te downloaden. Het moet niet mogelijk zijn om naar kantoor te komen, website te gebruiken of telefonisch contact op te nemen voor het procedure. Dus moet de content erop aangepast worden en de klant leiden naar het nieuwe app"
                        "Bestudeer de nieuwe website goed en schrijf de content daarop gebaseerd de nieuwe werkwijze moet overal in de content voorkomen en het oude werkwijze moet helemaal herschreven woorden naar het nieuwe werkwijze."
                        "Behoud de originele HTML-structuur en de Alinea-indeling, inclusief alinea's (<p>), koppen (<h2>, <h3>, <h4>), en opsommingen (<ul>, <li>). "
                        "1. Behoud de originele hoeveelheid woorden per alinea. De herschreven tekst mag niet minder woorden per alinea bevatten dan het origineel. "
                        "2. Gebruik een enkele <h1> voor de eerste kop,\n"
                        "3. Gebruik <h2>, <h3>, enzovoort voor subkoppen afhankelijk van de hiërarchie.\n"
                        "4. Alle paragrafen moeten worden ingesloten in <p>-tags.\n"
                        f"5. Gebruik het focus keyword '{focus_keyword}' in de koptekst (bijvoorbeeld <h1>) en integreer het natuurlijk in de eerste paragraaf. "
                        f"6. Zorg ervoor dat het focus keyword '{focus_keyword}' maximaal vijf keer voorkomt, verspreid over de tekst.\n\n"
                        
                        "Zorg ervoor dat de tekst leesbaar blijft en dat het focus keyword op een natuurlijke manier wordt geïntegreerd. "
                        "Vermijd overmatige herhaling van het keyword. Zorg ervoor dat de tekst goed geoptimaliseerd is voor zoekmachines, maar de leesbaarheid voor de gebruiker blijft behouden. "
                        "Pas de nieuwe werkwijze toe waar relevant, waarbij klanten worden aangemoedigd om de app te downloaden voor het aanmelden van auto's. "
                        "Maak gebruik van <h2> en <h3> voor belangrijke secties, en <p> voor paragrafen."
                        "herschrijf de inhoud zodanig dat het de nieuwe stappen weerspiegelt maximaal 1 keer in de content:"
                                "\n1. Meld uw auto aan via de app."
                                "\n2. Ontvang biedingen van gekwalificeerde dealers."
                                "\n3. Kies het beste bod en accepteer het."
                                "\n4. Rond de verkoop veilig af via de app."
                        "Gebruik specifieke Call-to-Actions (CTA's) zoals 'Download de app' en leg de nadruk op gebruiksgemak en snelheid van het proces. "
                        "Er mag geen tekst staan dat we een nieuwe werkwijze hebben alleen de herschrijving naar de nieuwe werkwijze"
                        "Ik verwacht geen enkel tekst over de oude werkwijze. bekijk elke pagina goed en implementeer de nieuwe werkwijze op de jusite manier."
                        f"{combined_content}"
                    )}
                ]

                response = client.chat.completions.create(
                    model="gpt-3.5-turbo",
                    messages=messages,
                    max_tokens=4000,
                    temperature=0.7,
                )
                rewritten_text = response.choices[0].message.content.strip()
                rewritten_parts.append(rewritten_text)

            rewritten_content = ' '.join(rewritten_parts)
            return rewritten_content

        except Exception as e:
            print(f"Error rewriting content: {e}")
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
    
    # Content herschrijven en focus keyword toevoegen
    original_content = row[column_content]
    rewritten_content = rewrite_content(original_content, focus_keyword)
    df.at[idx, column_content] = rewritten_content  

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
output_file = '../herschreven_excel/autoinkoop/auto5.xlsx'
df.to_excel(output_file, index=False)
print(f"Verwerking voltooid! De herschreven data is opgeslagen in {output_file}")
