## Bkn Living Content & Meta Generator

Een Python-applicatie voor het herschrijven van productinhoud en het genereren van SEO-geoptimaliseerde metadata voor de website van Bkn Living. Deze applicatie maakt gebruik van de OpenAI GPT-3.5 API om bestaande productinformatie te verbeteren en meta titles en descriptions aan te passen aan de laatste SEO-richtlijnen.

## Developers

Kerem Yildiz

## Projectoverzicht

Deze applicatie:

Herschrijft bestaande productinformatie voor Bkn Living-webpagina’s, waarbij de nadruk ligt op nauwkeurige en aantrekkelijke beschrijvingen.
Genereert SEO-geoptimaliseerde meta titles en descriptions die voldoen aan strikte SEO-eisen.
Verwijdert ongewenste karakters zoals aanhalingstekens en speciale symbolen uit de titels en beschrijvingen.
Verbetert de leesbaarheid door correcte Nederlandse termen en grammatica toe te passen.
Maakt gebruik van een krachtige keyword-optimalisatie om zoekwoorden nauwkeurig en natuurlijk in de content te verwerken.

## Installatie

Volg deze stappen om de applicatie lokaal op te zetten:

Clone de repository:

- git clone <repository-url>

Installeer de vereisten:

- pip install -r requirements.txt

OpenAI API-sleutel instellen:

- Maak een .env-bestand aan in de projectdirectory en voeg je OpenAI API-sleutel toe:
- OPENAI_API_KEY=your_openai_api_key_here

Plaats het Excel-bestand:

- Voeg het bknliving.xlsx bestand toe aan de map excel/bkn-living/.

## Gebruik

Start/ run de applicatie:

- python bkn.py

Werking van de applicatie:

De applicatie leest het bknliving.xlsx-bestand in, herschrijft de productbeschrijvingen en genereert SEO-geoptimaliseerde meta titles en descriptions. De herschreven content en metadata worden opgeslagen in een nieuw Excel-bestand in de map herschreven_excel/bkn-living/ met de naam updated_bkn-living.xlsx.

## Functionaliteiten

### 1. Focus Keywords Genereren

- Functie: improve_or_generate_focus_keyword
- De functie genereert of verbetert de focus keywords op basis van het producttype, kleur of materiaal en maakt deze geschikt voor SEO-doeleinden.

### 2. Producttitel Herschrijven

- Functie: rewrite_product_title
- Herschrijft producttitels tot een maximale lengte van 60 karakters, waarbij de focus keyword en de belangrijkste specificaties (zoals kleur en materiaal) worden benadrukt.

### 3. Productinhoud Herschrijven

- Functie: rewrite_product_content
- Maakt gebruik van het focus keyword om een aantrekkelijke productbeschrijving te schrijven die de belangrijkste productkenmerken en voordelen duidelijk weergeeft.

### 4. Meta Title Genereren

- Functie: generate_meta_title
- De functie maakt een korte en krachtige SEO-meta title van maximaal 50 karakters, inclusief focus keywords, zonder ongewenste tekens.

### 5. Meta Description Genereren

- Functie: generate_meta_description
- Creëert een aantrekkelijke meta description van maximaal 150 karakters, waarbij de belangrijkste productkenmerken in één zin worden weergegeven en een call-to-action wordt toegevoegd.

### 6. Global Unique ID Kopiëren

- Functie: copy_ean_to_global_unique_id
- Kopieert de waarde van de kolom meta:\_alg_ean naar meta:\_global_unique_id als een unieke productreferentie.

### 7. Alt-tekst voor Afbeeldingen Toevoegen

- Functie: add_alt_text_to_images
- Voegt alt-teksten toe aan de eerste afbeelding in de images-kolom. De alt-teksten zijn relevant, kort en bevatten het focus keyword voor optimale SEO.

## Resultaten

De herschreven content, meta titles en descriptions worden opgeslagen in een nieuw Excel-bestand (updated_bkn-living.xlsx) in de map herschreven_excel/bkn-living/.

# Resources

- OpenAI API
- Pandas Documentation

## Versie

- Versie: 1.0
- Release date: -

## Bugs

Geen bekende bugs.

## To-do's / Uitbreidingen

- Uitbreiding van foutafhandeling voor betere stabiliteit tijdens de inhoudsgeneratie.
- Mogelijkheid om handmatig titels en beschrijvingen aan te passen via een GUI voor nauwkeurige SEO-controle.
