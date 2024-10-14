## AutoInkoop Content & Meta Generator

Een Python-applicatie voor het herschrijven van webpagina-inhoud en het genereren van SEO-geoptimaliseerde meta titles en descriptions voor een auto-inkoopservice. De applicatie gebruikt de OpenAI GPT-3.5 API om bestaande content te herschrijven en meta-data te verbeteren volgens de nieuwste SEO-richtlijnen.

## Developers

Kerem Yildiz (1034004)

## Projectoverzicht

Deze applicatie:

- Herschrijft bestaande webinhoud voor autoinkoop-webpagina's, waarbij de nadruk ligt op de nieuwe werkwijze van AutoInkoopService.nl.
- Genereert SEO-geoptimaliseerde meta titles en descriptions die voldoen aan strikte SEO-regels.
- Verwijdert ongewenste tekens zoals aanhalingstekens en speciale symbolen uit de titels en beschrijvingen.
- Voegt automatisch de bedrijfsnaam toe als de meta title korter is dan 43 karakters.
- Optimaliseert meta descriptions om maximaal twee zinnen te bevatten, zonder dat de tekst wordt afgekapt.

## Installatie

Volg deze stappen om de applicatie lokaal op te zetten:

Clone de repository:

- git clone <repository-url>

Installeer de vereisten:

- pip install -r requirements.txt
- OpenAI API-sleutel instellen:

Maak een .env-bestand aan in de projectdirectory en voeg je OpenAI API-sleutel toe.

- OPENAI_API_KEY=your_openai_api_key_here

Plaats het Excel-bestand:

- Voeg het autoinkoop.xlsx bestand toe aan de map excel/.

## Gebruik

Start de applicatie:

- python main.py

## Werking van de applicatie:

De applicatie leest het autoinkoop.xlsx bestand in, herschrijft de webpagina-inhoud volgens de nieuwe werkwijze, en genereert SEO-geoptimaliseerde meta titles en descriptions.
De herschreven content, titels en beschrijvingen worden opgeslagen in een nieuw Excel-bestand in de map herschreven_excel/ met de naam herschreven_autoinkoop.xlsx.

## Functionaliteiten

1. Content Herschrijven
   De functie rewrite_content herschrijft de originele content door de nieuwe stappen van het verkoopproces van AutoInkoopService toe te passen:
   - Meld uw auto aan via de app.
   - Ontvang biedingen van gekwalificeerde dealers.
   - Kies het beste bod en accepteer het.
   - Rond de verkoop veilig af via de app.
   - HTML-structuur zoals koppen en opsommingen blijven intact.
   - Call-to-Actions (CTA's) worden toegevoegd om gebruikers aan te moedigen de app te downloaden.
2. Meta Title Genereren
   - De functie generate_meta_title genereert een korte en krachtige SEO-titel van maximaal 60 karakters.
   - Speciale tekens en aanhalingstekens worden verwijderd.
   - De bedrijfsnaam AutoInkoopService wordt toegevoegd als de titel minder dan 43 karakters bevat.
3. Meta Description Genereren
   - De functie generate_meta_description creëert een professionele en aantrekkelijke meta description van maximaal 150 karakters.
   - De description bestaat uit maximaal twee zinnen: de eerste geeft de belangrijkste boodschap weer, en de tweede bevat een call-to-action zonder ongewenste termen.
   - Speciale tekens en aanhalingstekens worden verwijderd, en afkappingen worden voorkomen.

## Resultaten

De herschreven content, meta titles en descriptions worden opgeslagen in een nieuw Excel-bestand (herschreven_autoinkoop.xlsx) in de map herschreven_excel/.

## Resources

- OpenAI API
- Pandas Documentation
- Flask Documentation

## Versie

- Versie 1.0
- Release date: -

## Bugs

- Geen bekende bugs.

## To-do's / Uitbreidingen

- Meer foutafhandelingsmechanismen toevoegen om beter om te gaan met uitzonderingen tijdens het herschrijven van de content.
- Mogelijkheid om titels en descriptions voor specifieke pagina's handmatig aan te passen via een gebruikersinterface.
