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
        "De title moet maximaal 60 karakters bevatten inclusief spaties, inclusief eventuele toevoegingen. "
        "Ik verwacht 1 korte en krachtige afgeronde titel van maximaal 60 karakter inclusief spaties."
        "Vermijd afgebroken zinnen en zorg ervoor dat het professioneel en aantrekkelijk is voor de lezer."
        "verwijder dit soort tekens '%%title%%','%%page%%', '%%sep%%'."
    )
    
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=100,
    )
    
    title = response['choices'][0]['message']['content'].strip()

    # Voeg bedrijfsnaam toe als de titel minder dan 43 karakters is
    if len(title) < 43:
        full_title = f"{title} | {company_name}"
    else:
        full_title = title

    # Controleer de lengte van de uiteindelijke titel
    if len(full_title) > 60:
        # Bereken het maximale aantal karakters voor de titel
        max_title_length = 60 - len(f" | {company_name}")  # 60 - lengte van de bedrijfsnaam
        title = title[:max_title_length].rsplit(' ', 1)[0]  # Knip de titel af op de laatste spatie
        full_title = f"{title} | {company_name}"

    return full_title

def generate_meta_description(subject, focus_keyword, existing_description):
    prompt = (
        f"Schrijf een SEO-geoptimaliseerde meta description voor een informatieve pagina over {subject}. "
        f"Het zoekwoord is '{focus_keyword}'. "
        "Ik wil maximaal 2 zinnen in de description (dit zijn strikte regels): "
        "De eerste zin moet de belangrijkste boodschap van de pagina duidelijk maken. "
        "De tweede zin moet een duidelijke call-to-action bevatten. "
        "Gebruik geen termen zoals 'vraag om een bod'. Vervang het met 'ontvang een bod'. "
        "Gebruik geen termen zoals 'ontvang cash'."
        "verwijder dit soort tekens '%%title%%','%%page%%','%%sep%%' ."
        "Zorg ervoor dat de beschrijving maximaal 150 karakters bevat inclusief spaties (dit zijn strikte regels) en ik wil geen afkappingen. het moet een afgerond beschrijving zijn "
        "Maak het alogisch, professioneel en aantrekkelijk voor de lezer. "
        f"Hier is een bestaande meta description die je kan gebruiken voor logica: {existing_description}."
    )
    
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=150,
    )
    
    description = response['choices'][0]['message']['content'].strip()

    # Check of de description meer dan 150 karakters is
    if len(description) > 150:
        # Fallback: inkorten van de description om de limiet te respecteren
        description = description[:147].rsplit(' ', 1)[0] + '...'
    
    # Check of het aantal zinnen meer dan 2 is
    if description.count('.') > 2:
        sentences = description.split('.')
        description = '. '.join(sentences[:2]) + '.'

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