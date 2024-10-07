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

def rewrite_content(content):
    if pd.notnull(content):
        try:
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "user", "content": (
                            "Herschrijf de volgende webpagina-inhoud naar de nieuwe werkwijze van new.autoInkoopService.nl. "
                            "De nieuwe werkwijze gebruikt een app voor het aanmelden van auto's, en het doel is om klanten aan te moedigen de app te downloaden. "
                            "Maak de inhoud korter, maar zorg ervoor dat deze informatief en klantgericht blijft. "
                            "Behoud de HTML-structuur (zoals <h2>, <p>, <ul>, etc.), en zorg ervoor dat de nieuwe inhoud de voordelen van het platform en de app benadrukt. "
                            "Verander de inhoud zodanig dat het de nieuwe stappen weerspiegelt:"
                            "\n1. Meld uw auto aan via de app."
                            "\n2. Ontvang biedingen van gekwalificeerde dealers."
                            "\n3. Kies het beste bod en accepteer het."
                            "\n4. Rond de verkoop veilig af via de app."
                            "Gebruik specifieke CTA's (Call-to-Actions) zoals 'Download de app' en zorg ervoor dat de focus ligt op het gebruiksgemak en de snelheid van het proces."
                            "Hieronder is de originele content:"
                            f"\n\n{content}"
                        )
                    }
                ],
                max_tokens=1000,
            )
            rewritten_text = response['choices'][0]['message']['content']
            return rewritten_text.strip()
        except Exception as e:
            print(f"Error rewriting content: {e}")
            return content  # Retourneer originele tekst bij een fout
    return content

# Vervang de inhoud in de originele kolom 'Content'
df[column_content] = df[column_content].apply(rewrite_content)

# Sla de gewijzigde DataFrame op in hetzelfde Excel-bestand of een nieuw bestand
df.to_excel('excel/herschreven_test_10_pagina.xlsx', index=False)

print("De content is herschreven en opgeslagen.")