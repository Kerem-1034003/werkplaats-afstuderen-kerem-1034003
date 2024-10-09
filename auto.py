import os
from dotenv import load_dotenv
import pandas as pd
import openai

# OpenAI API key laden
load_dotenv()
openai_api_key = os.getenv('OPENAI_API_KEY')
openai.api_key = openai_api_key

# Excel bestand inlezen
df = pd.read_excel('excel/autoinkoop.xlsx')
column_content = 'Content'
column_meta_title = '_yoast_wpseo_title'
column_meta_description = '_yoast_wpseo_metadesc'
column_focus_keyword = '_yoast_wpseo_focuskw'

max_title_length = 60  
max_description_length = 160  
company_name = "AutoInkoopService"


# Functie om tekst te splitsen op basis van alinea's
def split_text_by_paragraphs(text, max_paragraphs):
    paragraphs = text.split("\n\n")  # Alinea's splitsen op dubbele enters
    return ['\n\n'.join(paragraphs[i:i + max_paragraphs]) for i in range(0, len(paragraphs), max_paragraphs)]

# Functie om content te herschrijven met dynamisch splitsen en contextbehoud
def rewrite_content(content):
    if pd.notnull(content):
        try:
            # Tellen van het aantal woorden
            word_count = len(content.split())

            # Als het aantal woorden meer dan 1500 is, splitsen we de tekst
            if word_count > 1000:
                paragraphs = content.split("\n\n")
                split_contents = split_text_by_paragraphs(content, max_paragraphs=6)  # 6 alinea's per deel
            else:
                split_contents = [content]  # Geen splitsing nodig voor kortere teksten

            rewritten_parts = []
            for idx, part in enumerate(split_contents):
                # Context behouden met vorige deel (niet te groot maken)
                context_text = "\n\n".join(rewritten_parts[-1:]) if rewritten_parts else ""
                combined_content = f"{context_text}\n\n{part}"

                response = openai.ChatCompletion.create(
                    model="gpt-3.5-turbo",
                    messages=[
                        {"role": "user", "content": (
                                "Herschrijf de volgende webpagina-inhoud naar de nieuwe werkwijze van new.autoInkoopService.nl en behoud de alinea indeling. "
                                "Behoud de originele HTML-structuur, inclusief alinea's (<p>), koppen (<h2>, <h3>, <h4>), en opsommingen (<ul>, <li>). "
                                "Pas de nieuwe werkwijze toe waar relevant, waarbij klanten worden aangemoedigd om de app te downloaden voor het aanmelden van auto's. "
                                "Zorg ervoor dat elk alinea inhoud gedetailleerd, informatief en klantgericht is. "
                                "Gebruik de app als een logische toevoeging binnen het verkoopproces, maar vermijd herhaling."
                                "Geef voorbeelden en verduidelijkingen waar mogelijk."
                                "Hier is de originele content, behoud de alinea-indeling exact zoals deze is:"
                                "Verander de inhoud zodanig dat het de nieuwe stappen weerspiegelt:"
                                "\n1. Meld uw auto aan via de app."
                                "\n2. Ontvang biedingen van gekwalificeerde dealers."
                                "\n3. Kies het beste bod en accepteer het."
                                "\n4. Rond de verkoop veilig af via de app."
                                "Gebruik specifieke Call-to-Actions (CTA's) zoals 'Download de app' en leg de nadruk op gebruiksgemak en snelheid van het proces. "
                                "Hieronder is de originele content:"
                                f"\n\n{combined_content}"
                            )}
                            ],
                        max_tokens=4096,  # Gebruik volledige limiet
                    )
                rewritten_text = response['choices'][0]['message']['content']
                rewritten_parts.append(rewritten_text.strip())  # Verzamel elk herschreven deel

            # Voeg de herschreven delen samen
            return ' '.join(rewritten_parts)
        
        except Exception as e:
            print(f"Error rewriting content: {e}")
            return content  # Retourneer originele tekst bij een fout
    return content

# Functie om nieuwe meta title te genereren
def generate_meta_title(subject, focus_keyword):
    prompt = (
        f"Schrijf een SEO-geoptimaliseerde meta title voor een informatieve pagina over {subject}. "
        f"Het zoekwoord is '{focus_keyword}'. "
        "De title moet maximaal 60 karakters bevatten en eindigen met een logische, afgeronde zin. "
        "Vermijd afgebroken zinnen en zorg ervoor dat het professioneel en aantrekkelijk is voor de lezer. "
        "Indien van toepassing, eindig met ' | AutoInkoopService'."
    )
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=100,
    )
    title = response['choices'][0]['message']['content'].strip()
    return title


def generate_meta_description(subject, focus_keyword, existing_description):
    prompt = (
        f"Schrijf een SEO-geoptimaliseerde meta description voor een informatieve pagina over {subject}. "
        f"Het zoekwoord is '{focus_keyword}'. "
        "De eerste zin moet de belangrijkste boodschap van de pagina duidelijk maken. "
        "Gebruik geen termen zoals 'vraag om een bod'. "
        "De tweede zin moet de voordelen van het gebruik van de app benadrukken zonder afgebroken zinnen. "
        "De derde zin moet een duidelijke call-to-action bevatten. "
        "Zorg ervoor dat de beschrijving maximaal 150 karakters bevat en ik wil geen afkappingen. "
        "Maak het logisch, professioneel en aantrekkelijk voor de lezer. "
        f"Hier is een bestaande meta description: {existing_description}."
    )
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=200,
    )
    description = response['choices'][0]['message']['content'].strip()
    return description

new_meta_titles = []
new_meta_descriptions = []

for idx, row in df.iterrows():
    # Herschrijf de content
    original_content = row[column_content]
    df.at[idx, column_content] = rewrite_content(original_content)

    # Genereer nieuwe meta title en description
    subject = row[column_meta_title] if pd.notnull(row[column_meta_title]) else "Geen onderwerp"
    focus_keyword = row[column_focus_keyword] if pd.notnull(row[column_focus_keyword]) else "geen zoekwoord"
    existing_description = row[column_meta_description] if pd.notnull(row[column_meta_description]) else ""

    new_meta_title = generate_meta_title(subject, focus_keyword)
    new_meta_description = generate_meta_description(subject, focus_keyword, existing_description)

    new_meta_titles.append(new_meta_title)
    new_meta_descriptions.append(new_meta_description)

# Voeg de nieuwe meta titles en descriptions toe aan de DataFrame
df[column_meta_title] = new_meta_titles
df[column_meta_description] = new_meta_descriptions

# Sla de gewijzigde DataFrame op in een nieuw Excel-bestand
df.to_excel('excel/herschreven_compleet.xlsx', index=False)

print("De content, meta titles en descriptions zijn herschreven en opgeslagen.")