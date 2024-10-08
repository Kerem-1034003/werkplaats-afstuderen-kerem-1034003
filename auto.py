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
                                "Herschrijf de volgende webpagina-inhoud naar de nieuwe werkwijze van new.autoInkoopService.nl. "
                                "Behoud de originele HTML-structuur, inclusief alinea's (<p>), koppen (<h2>, <h3>, <h4>), en opsommingen (<ul>, <li>). "
                                "Pas de nieuwe werkwijze toe waar relevant, waarbij klanten worden aangemoedigd om de app te downloaden voor het aanmelden van auto's. "
                                "Zorg ervoor dat de inhoud gedetailleerd, informatief en klantgericht is. "
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

# Vervang de inhoud in de originele kolom 'Content'
df[column_content] = df[column_content].apply(rewrite_content)

# Sla de gewijzigde DataFrame op in hetzelfde Excel-bestand of een nieuw bestand
df.to_excel('excel/herschreven_test_10_pagina.xlsx', index=False)

print("De content is herschreven en opgeslagen.")
