import os
from dotenv import load_dotenv
import pandas as pd
from openai import OpenAI
import time

# OpenAI API key laden
load_dotenv()
openai_api_key = os.getenv('OPENAI_API_KEY')
client = OpenAI(api_key=openai_api_key)

# Excel bestand inlezen

df = pd.read_excel('excel/autoinkoop/autoinkoop_split_final_6.xlsx')

column_content = 'Content'
column_meta_title = '_yoast_wpseo_title'
column_meta_description = '_yoast_wpseo_metadesc'
column_focus_keyword = '_yoast_wpseo_focuskw'

max_title_length = 60  
max_description_length = 150  
company_name = "AutoInkoopService"

# Functie om tekst te splitsen op basis van alinea's
def split_text_by_paragraphs(text, max_paragraphs):
    paragraphs = text.split("\n\n")
    return ['\n\n'.join(paragraphs[i:i + max_paragraphs]) for i in range(0, len(paragraphs), max_paragraphs)]

# Functie om content te herschrijven met dynamisch splitsen en contextbehoud
def rewrite_content(content):
    if pd.notnull(content):
        try:
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
                        "Herschrijf de volgende webpagina-inhoud voor de nieuwe werkwijze van AutoInkoopService.nl, "
                        "inclusief nieuwe app-stappen en verbeterde klantenservice." 
                        "Behoud de structuur en voeg CTA's toe voor app-downloads."
                        f"\n\n{combined_content}"
                    )}
                ]

                response = client.chat.completions.create(
                    model="gpt-3.5-turbo",
                    messages=messages,
                    max_tokens=2000,
                    temperature=0.7
                )
                rewritten_text = response.choices[0].message.content.strip()
                rewritten_parts.append(rewritten_text)

            return ' '.join(rewritten_parts)

        except Exception as e:
            print(f"Error rewriting content: {e}")
            return content
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
                    max_tokens=1000,
                    temperature=0.7
                )
                additional_text = response.choices[0].message.content.strip()
                content += " " + additional_text
            return content

        except Exception as e:
            print(f"Error filling content: {e}")
            return content
    return content

# Functie om nieuwe meta title te genereren
def generate_meta_title(subject, focus_keyword):
    prompt = (
        f"Schrijf een SEO-geoptimaliseerde meta title voor {subject}, "
        f"met het zoekwoord '{focus_keyword}', maximaal 60 karakters. "
        f"Bijvoorbeeld: 'Snel uw auto verkopen | {company_name}'."
    )
    try:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=60,
            temperature=0.6
        )
        title = response.choices[0].message.content.strip()
        title = title.replace('"', '').replace("'", "")
        return f"{title} | {company_name}" if len(title) < 43 else title
    except Exception as e:
        print(f"Error generating meta title: {e}")
        return ""

# Functie om nieuwe meta description te genereren
def generate_meta_description(subject, focus_keyword, existing_description):
    prompt = (
        f"Schrijf een SEO-geoptimaliseerde meta description over {subject}, "
        f"gebruik het zoekwoord '{focus_keyword}', en houd het binnen 150 karakters, inclusief CTA."
    )
    try:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=100,
            temperature=0.6
        )
        description = response.choices[0].message.content.strip()
        return description[:150] if len(description) > 150 else description
    except Exception as e:
        print(f"Error generating meta description: {e}")
        return existing_description

new_meta_titles = []
new_meta_descriptions = []

for idx, row in df.iterrows():
    print(f"Processing row {idx + 1} of {len(df)}")
    original_content = row[column_content]
    rewritten_content = rewrite_content(original_content)
    filled_content = fill_content(rewritten_content)
    df.at[idx, column_content] = filled_content

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

output_file = 'herschreven_excel/autoinkoop/herschreven_autoinkoop_split_final_6.xlsx'
df.to_excel(output_file, index=False)

print("De content, meta titles en descriptions zijn herschreven en opgeslagen.")
