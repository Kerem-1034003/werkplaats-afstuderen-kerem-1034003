import os
from dotenv import load_dotenv
import pandas as pd
import openai

# OpenAI API key laden
load_dotenv()
openai_api_key = os.getenv('OPENAI_API_KEY')
openai.api_key = openai_api_key

df = pd.read_excel('excel/bnb-living/bnbpp.xlsx')

column_slug = 'slug'
column_seo_title = 'seo_title'
column_seo_description = 'seo_description'
column_focus_keyword = 'focus_keyword'

company_name = "BnB Living"

# Functie om een focus keyword te genereren op basis van de slug
def generate_focus_keyword(slug):
    prompt = (
        f"Op basis van de slug '{slug}', genereer een relevant focus keyword dat kort, specifiek en SEO-geoptimaliseerd is. "
        "Houd bij het genereren van het keyword rekening met de zoekintentie van de gebruiker en gebruik indien mogelijk semantische zoekwoorden of gerelateerde termen om de context te versterken."
    )
    
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=30,
    )
    
    focus_keyword = response['choices'][0]['message']['content'].strip()

    focus_keyword = focus_keyword.replace('"', '').replace("'", "")

    return focus_keyword

# Functie om een SEO-titel te genereren op basis van het focus keyword
def generate_seo_title(focus_keyword):
    prompt = (
        f"Schrijf een krachtige, korte SEO-geoptimaliseerde meta title met het focus keyword '{focus_keyword}'. "
        "De titel moet aantrekkelijk zijn en binnen 5-6 woorden passen, zonder langer te zijn dan 60 karakters inclusief spaties. "
        "Zorg ervoor dat de titel goed aansluit op de zoekintentie van de gebruiker en dat deze niet wordt afgebroken. "
        "Houd het professioneel, pakkend en relevant voor de content van de pagina."
    )

    
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=60,
    )
    
    title = response['choices'][0]['message']['content'].strip()

    title = title.replace('"', '').replace("'", "")

    # Voeg de bedrijfsnaam toe, als de titel minder dan 43 karakters is
    if len(title) < 43:
        full_title = f"{title} | {company_name}"
    else:
        full_title = title

    return full_title

# Functie om een SEO-beschrijving te genereren op basis van het focus keyword
def generate_seo_description(focus_keyword):
    prompt = (
        f"Schrijf een SEO-geoptimaliseerde meta description die volledig aansluit bij de zoekintentie van de gebruiker voor het focus keyword '{focus_keyword}'. "
        "De beschrijving moet kort en krachtig zijn, bestaan uit maximaal 2 zinnen, waarin de eerste zin de belangrijkste boodschap van de pagina helder samenvat, "
        "en de tweede zin een duidelijke call-to-action bevat die aansluit op de content. Gebruik waar mogelijk semantische varianten van het focus keyword voor context. "
        "De beschrijving mag niet langer zijn dan 150 karakters inclusief spaties. Zorg ervoor dat de beschrijving logisch, professioneel en aantrekkelijk is voor de lezer."
    )
    
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=150,
    )
    
    description = response['choices'][0]['message']['content'].strip()

     # Verwijder ongewenste tekens
    description = description.replace('"', '').replace("'", "")

    # Opsplitsen in zinnen
    sentences = description.split('. ')
    
    # Houd maximaal twee zinnen
    if len(sentences) > 2:
        description = '. '.join(sentences[:2]) + '.'

    # Check of de lengte meer dan 150 karakters is
    if len(description) > 150:
        # Knip de tekst af na de laatste volledige zin binnen de limiet
        shortened_description = ''
        for sentence in sentences:
            if len(shortened_description) + len(sentence) + 1 <= 150:  # +1 voor de punt
                shortened_description += sentence + '. '
            else:
                break
        
        # Verwijder de laatste spatie en punt als de lengte te kort is
        description = shortened_description.strip()

    return description

# Nieuwe kolommen voor SEO-titels, beschrijvingen en focus keywords
new_seo_titles = []
new_seo_descriptions = []
new_focus_keywords = []

for idx, row in df.iterrows():
    slug = row[column_slug]
    
    # Genereer focus keyword
    focus_keyword = generate_focus_keyword(slug)
    
    # Genereer SEO titel en beschrijving
    seo_title = generate_seo_title(focus_keyword)
    seo_description = generate_seo_description(focus_keyword)
    
    # Sla resultaten op in lijsten
    new_focus_keywords.append(focus_keyword)
    new_seo_titles.append(seo_title)
    new_seo_descriptions.append(seo_description)

# Voeg de nieuwe gegevens toe aan de DataFrame
df[column_focus_keyword] = new_focus_keywords
df[column_seo_title] = new_seo_titles
df[column_seo_description] = new_seo_descriptions

# Sla de gewijzigde DataFrame op in een nieuw Excel-bestand
df.to_excel('herschreven_excel/bnb-living/herschreven_bnbliving.xlsx', index=False)

print("De SEO-titels, descriptions en focus keywords zijn gegenereerd en opgeslagen.")

