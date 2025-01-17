import pandas as pd
import json
import os
from openai import OpenAI
from dotenv import load_dotenv
import re

# Laad API-sleutels
load_dotenv()
openai_api_key = os.getenv('OPENAI_API_KEY')

# Initialiseer de OpenAI client
client = OpenAI(api_key=openai_api_key)

# Paden voor input en output bestanden
input_folder = '../herschreven_excel/simpledeal/homcomsplit/'  # Map met input Excel-bestanden
output_folder = '../herschreven_excel/simpledeal/homcomimport/'  # Map voor output-bestanden
json_path = '../v10_datamodel_v10_nl.json'
os.makedirs(output_folder, exist_ok=True)  # Zorg ervoor dat de outputmap bestaat

# Mapping-tabel: Excel-categorieën naar JSON-categorieën
category_mapping = {
    'Wonen>Elektronica>Elektrische haarden': 'Sfeerhaard',
    'Wonen>Elektronica>Elektrische Kachels': 'Kachel',
    'Wonen>Elektronica>Ethanol haarden': 'Sfeerhaard',
    'Wonen>Elektronica>Open haard accessoires': 'Kachel accessoire',

    'Wonen>Banken>Bevenbanken': 'Bed',
    'Wonen>Banken>Hockers': 'Hocker',
    'Wonen>Banken>Hoekbanken': 'Bank',
    'Wonen>Banken>Schoenenbanken': 'Schoenenkast',
    'Wonen>Banken>Slaapbank': 'Bed',
    'Wonen>Banken>Zitbank': 'Bank',

    'Wonen>Bedden>Gastbed': 'Bed',
    'Wonen>Bedden>Stapelbed': 'Bed',

    'Wonen>Kasten>Archiefkast': 'Kast',
    'Wonen>Kasten>Keukenkasten': 'Keukenkast',
    'Wonen>Kasten>Medicijnkastje': 'Badkamerkast',
    'Wonen>Kasten>Opbergkasten': 'Kast',
    'Wonen>Kasten>Schoenenkast': 'Schoenenkas',
    'Wonen>Kasten>Boekenkast': 'Kast',
    'Wonen>Kasten>Kledingkast': 'Kast',
    'Wonen>Kasten>Nachtkastjes': 'Kast',
    'Wonen>Kasten>Sieradenkasten': 'Sieradenopberger',
    'Wonen>Kasten>Stellingkasten': 'Kast',
    'Wonen>Kasten>TV-meubels': 'Kast',
    'Wonen>Kasten>Spiegelkasten': 'Badkamerkast',
    'Wonen>Kasten>Badkamerkast': 'Badkamerkast',

    'Wonen>Stoelen>Bureaustoelen': 'Bureaustoel',
    'Wonen>Stoelen>Eetkamerstoelen': 'Eetkamerstoel',
    'Wonen>Stoelen>Fauteuil': 'Fauteuil',
    'Wonen>Stoelen>Gamingstoel': 'Bureaustoel',
    'Wonen>Stoelen>Hangstoel': 'Hangstoel',
    'Wonen>stoelen>Massagestoelen': 'Fauteuil',
    'Wonen>Stoelen>Schommelstoel': 'Schommelstoel',
    'Wonen>Stoelen>Stoelkrukken': 'Krukje',
    'Wonen>Stoelen>Barkrukken': 'Barkruk',
    'Wonen>Stoelen>Zitballen': 'Zitbal',

    'Wonen>Tafels>Bartisch set': 'Tafel',
    'Wonen>Tafels>Bartafels': 'Tafel',
    'Wonen>Tafels>Bijzettafels': 'Tafel',
    'Wonen>Tafels>Eettafels': 'Tafel',
    'Wonen>Tafels>Bureaus': 'Bureau',
    'Wonen>Tafels>Eethoeken': 'Set tafel en stoelen',
    'Wonen>Tafels>Hoekstafels': 'Tafel',
    'Wonen>Tafels>Consoletafels': 'Tafel',
    'Wonen>Tafels>Koffie en Bijzettafels': 'Tafel',
    'Wonen>Tafels>Klaptafel': 'Tafel',
    'Wonen>Tafels>Massagetafels': 'Massagetafel',
    'Wonen>Tafels>Salontafels': 'Tafel',
    'Wonen>Tafels>Serveerwagen': 'Keukentrolley',

    'Wonen>Verlichting>Plafondlampen': 'Plafonnière',
    'Wonen>Verlichting>Staandelampen': 'Staande lamp',
    'Wonen>Verlichting>Tafellampen': 'Tafellamp',

    'Wonen>Planken & Rekken>Badkamerplank': 'Badkamerrek',
    'Wonen>Planken & Rekken>Boekenplank': 'Boekenplank',
    'Wonen>Planken & Rekken>Schoenenrek': 'Schoenenopbergers',
    'Wonen>Planken & Rekken>Wandplanken': 'Wandrek',

    'Wonen>Accessoires>Brievenbussen': 'Brievenbus',
    'Wonen>Accessoires>Toilettassen': 'Toilettas',
    'Wonen>Accessoires>Fotolijsten': 'Fotolijst',
    'Wonen>Accessoires>Horlogeboxen': 'Horlogedoos',
    'Wonen>Accessoires>Sieradendozen': 'Sieradenopberger',
    'Wonen>Accessoires>Make-up organizers': 'Make-up organizer',
    'Wonen>Accessoires>Make-up Spiegel': 'Make-upspiegel',
    'Wonen>Accessoires>Vloerkleden': 'Vloerkleed',
    'Wonen>Accessoires>Paraplubak': 'Paraplubakhouder',
    'Wonen>Accessoires>Wasmand': 'Wasmand',

    'Huishouden>Keuken>Keukenkrukken': 'Keukenkruk',
    'Huishouden>Keuken>Keukentrolley': 'Keukentrolley',
    'Huishouden>Keuken>Kruidenrek': 'Kruidenrek',
    'Huishouden>Keuken>Prullenbakken': 'Prullenbak',

    'Huishouden>Handgereedschap>Zaagbokken': 'Zaagbok',
    'Huishouden>Handgereedschap>Ladders': 'Ladder',
    'Huishouden>Handgereedschap>Gereedschapskoffers': 'Gereedschapskoffer',
    'Huishouden>Handgereedschap>Gereedschapskist': 'Gereedschapskist',
    'Huishouden>Handgereedschap>Gereedschapswanden': 'Gereedschapswand',
    'Huishouden>Handgereedschap>Gereedschapskarren': 'Gereedschapskar',
    'Huishouden>Handgereedschap>Metaaldetectors': 'Detectieapparaat',

    'Huishouden>Opbergen>Boodschapentrolley': 'Boodschapentrolley',
    'Huishouden>Opbergen>Kledinghangers': 'Kledinghanger',
    'Huishouden>Opbergen>Opbergdozen': 'Opberger',
    'Huishouden>Opbergen>Opbergkisten': 'Opberger',
    'Huishouden>Opbergen>Stellingkasten': 'Kast',

    'Huishouden>Rond het huis>Fittingen in de schuifdeur': 'Schuifdeuronderdeel',
    'Huishouden>Rond het huis>Luifels': 'Luifel',
    'Huishouden>Rond het huis>Kamerverdeler': 'Kamerscherm',
    'Huishouden>Rond het huis>Kunstplanten': 'Kunstplant',
    'Huishouden>Rond het huis>Kluizen': 'Kassakluis',
    'Huishouden>Rond het huis>Schuifdeuren': 'Deur',
    'Huishouden>Rond het huis>Rol Gordijnen': 'Rolgordijn',

    'Huishouden>Steekwagen>Boderkarren': 'Bolderkar',
    'Huishouden>Steekwagen>Steekwagens': 'Steekwagen',

    'Baby & Kind>Babybadkamer>Babybadjes': 'Babybadje',
    'Baby & Kind>Babybadkamer>Babytoilet': 'Plaspot',

    'Baby & Kind>Fietskar>Fietskar voor kinderen': 'Fietskar',

    'Baby & Kind>Buitenspeelgoed>Kinderglijbanen': 'Glijbaan',
    'Baby & Kind>Buitenspeelgoed>Kinderschommels': 'Schommel',
    'Baby & Kind>Buitenspeelgoed>Speelhuisjes & Tenten': 'Speelhuis',
    'Baby & Kind>Buitenspeelgoed>Springkussens en waterkastelen': 'Springkussen',
    'Baby & Kind>Buitenspeelgoed>Zandbakken & toebehoren': 'Zandbak',

    'Baby & Kind>Kindersporten>Basketbalstandaarts': 'Basketbalstandaard',
    'Baby & Kind>Kindersporten>Voetbaldoelen': 'Voetbaldoel',
    'Baby & Kind>Kindersporten>Balansstenen': 'Balansspeelgoed',
    'Baby & Kind>Kindersporten>Kindertrampoline': 'Trampoline',

    'Baby & Kind>Kindermeubels>Kinderzitmeubels': 'Bank',
    'Baby & Kind>Kindermeubels>Kinder Tafels': 'Tafel',
    'Baby & Kind>Kindermeubels>Kinderplanken & Kinderkasten': 'Kast',
    'Baby & Kind>Kindermeubels>Leertoren': 'Leertoren',

    'Baby & Kind>Kindervoertuigen>Driewielers': 'Driewieler',
    'Baby & Kind>Kindervoertuigen>Steps': 'Step',
    'Baby & Kind>Kindervoertuigen>Loopfietsen': 'Loopfiets',
    'Baby & Kind>Kindervoertuigen>Scooter': 'Scooter (elektrisch)',
    'Baby & Kind>Kindervoertuigen>Skelters': 'Skelter',

    'Baby & Kind>Speelgoed>Poppenhuizen': 'Poppenhuis',
    'Baby & Kind>Speelgoed>Puzzelmatten': 'Puzzelmat',
    'Baby & Kind>Speelgoed>Schommelende dieren': 'Hobbelfiguur',
    'Baby & Kind>Speelgoed>Speelbouwstenen': 'Constructiespeelgoed',

    'Huisdieren>Honden>Behendigheidspeelgoed': 'Speelgoed voor dieren',
    'Huisdieren>Honden>Hondenbed': 'Dierenmand',
    'Huisdieren>Honden>Hondenhokken': 'Hok',
    'Huisdieren>Honden>Hondenkooien': 'Autostoel voor dieren',
    'Huisdieren>Honden>Hondenzwembaden': 'Zwembad',
    'Huisdieren>Honden>Manden': 'Dierenmand',
    'Huisdieren>Honden>Draagtassen': 'Draagtas',
    'Huisdieren>Honden>Autostoelen voor honden': 'Autostoel voor dieren',
    'Huisdieren>Honden>Beveiling en behuizing van honden': 'Veiligheidshekje',
    'Huisdieren>Honden>Hondentrappen': 'Loopplank voor dieren',
    'Huisdieren>Honden>Hondenkar': 'Fietskar',
    'Huisdieren>Honden>Hondentraphekjes': 'Veiligheidshekje',
    'Huisdieren>Honden>Puppy Training Pads': 'Zindelijkheidstraining',
    'Huisdieren>Honden>Trimtafels': 'Trimtafel',
    'Huisdieren>Honden>Voedingskommen': 'Voerbak',

    'Huisdieren>Vogelaccessoires>Kippenhokken': 'Hok',
    'Huisdieren>Vogelaccessoires>Vogelkooien': 'Kooi',

    'Huisdieren>Katten>Katspeelgoed': 'Speelgoed voor dieren',
    'Huisdieren>Katten>Kattenbakken': 'Kattenbak',
    'Huisdieren>Katten>Krabpalen': 'Krabpaal',

    'Huisdieren>Knaagdierenkooien>Caviakooien': 'Kooi',
    'Huisdieren>Knaagdierenkooien>Konijnenkooien': 'Kooi',
    'Huisdieren>Knaagdierenkooien>Terrariums': 'Terrarium',

    'Sport>Sporten>Basketbal': 'Basketbal',
    'Sport>Sporten>Darten': 'Dartbord',
    'Sport>Sporten>Gymanstiek en Yoga': 'Yogabal',
    'Sport>Sporten>Trampolines en accesoires': 'Trampoline',
    'Sport>Sporten>Voetbaldoel': 'Voetbaldoel',
    'Sport>Sporten>Badminton': 'Badmintonracket',

    'Sport>Boksen>Boksbal': 'Boksbal',
    'Sport>Boksen>Bokszak': 'Bokszak',

    'Sport>Fitness & Krachtsport>Loopbanden': 'Loopband',
    'Sport>Fitness & Krachtsport>Dumbells': 'Dumbbells',
    'Sport>Fitness & Krachtsport>Halters': 'Dumbbells',
    'Sport>Fitness & Krachtsport>Fitnessmaterialen': 'Accessoires voor fitnessapparatuur',
    'Sport>Fitness & Krachtsport>Ergometer en home trainer': 'Hometrainer',
    'Sport>Fitness & Krachtsport>Aerobic step': 'Aerobic stepper',

    'Sport>Fiets accessoires>Fietsaanhangwagen': 'Fietskar',
    'Sport>Fiets accessoires>Helm': 'Sporthelm',
    'Sport>Fiets accessoires>Fietsmontagestandaard': 'Montagestandaard',
    'Sport>Fiets accessoires>Fietsbeugel': 'Fietsbeugel',

    'Tuin>Tuinschuren en tuinkassen>Tuinhuizen': 'Tuinhuis',
    'Tuin>Tuinschuren en tuinkassen>Tuinkassen': 'Kas',
    'Tuin>Tuinschuren en tuinkassen>Koude kassen': 'Moestuinbak',

    'Tuin>Pavilion Daken>Paviljoens en feesttenten': 'Partytent',
    'Tuin>Pavilion Daken>Pergola': 'Overkapping',

    'Tuin>Grillen en verwarmen>Straalkachels': 'Kachel',
    'Tuin>Grillen en verwarmen>Grill Pavilions': 'Partytent',
    'Tuin>Grillen en verwarmen>Grills': 'Barbecue',
    'Tuin>Grillen en verwarmen>Vuurschalen': 'Vuurschaal',

    'Tuin>Tuin accessoires>Tuinslangen': 'Tuinslang',
    'Tuin>Tuin accessoires>Tuinwagen': 'Tuinkar',
    'Tuin>Tuin accessoires>Tuincomposter': 'Potgrond',

    'Tuin>Tuindecoratie>Kunstgras': 'Kunstgras',
    'Tuin>Tuindecoratie>Tuin Bruggen': 'Loopbrug',
    'Tuin>Tuindecoratie>Plantenbak': 'Plantenbak',
    'Tuin>Tuindecoratie>Plant Tafels': 'Plantentafel',
    'Tuin>Tuindecoratie>Fontein': 'Fontein',
    'Tuin>Tuindecoratie>Tuinpoorten': 'Tuinpoort',
    'Tuin>Tuindecoratie>Windschermen': 'Tuinscherm',
    'Tuin>Tuindecoratie>Privacy Protection & Art Hedges': 'Tuinhek',
    'Tuin>Tuindecoratie>Tuinverlichting': 'Padverlichting',

    'Tuin>Zonwering>Parasol': 'Parasol',
    'Tuin>Zonwering>Parasol standaard': 'Parasolvoet',
    'Tuin>Zonwering>Parasol Hoes': 'Tuinmeubelhoes',
    'Tuin>Zonwering>Luifel': 'Luifel',

    'Tuin>Camping>Campingstoelen': 'Campingstoel',
    'Tuin>Camping>Campingbed': 'Campingbedje',
    'Tuin>Camping>Camping tafel': 'Campingtafel',
    'Tuin>Camping>Veldbedden': 'Campingbedje',
    'Tuin>Camping>Douchen en toiletkampen': 'Campingdouche'
}

# Functies

def load_json_file(path):
    try:
        with open(path, encoding='utf-8') as file:
            return json.load(file)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"Fout bij het laden van JSON-bestand: {e}")
        exit(1)

def get_required_attributes(category_name, products_data):
    for category, products in products_data.items():
        for product in products:
            if isinstance(product, dict) and product.get('name', '').lower() == category_name.lower():
                return [
                    {
                        'id': attr['id'],
                        'lovId': attr.get('lovId')
                    } 
                    for attr in product.get('attributes', []) 
                    if attr.get('enrichmentLevel') == 1
                ]
    return []

def get_attribute_values(lov_id, products_data):
    if not lov_id:
        return None
    for category, products in products_data.items():
        for product in products:
            if isinstance(product, dict) and product.get('id') == lov_id:
                return product.get('values', [])
    return None

def generate_attribute_value(post_content, attribute_name, allowed_values=None):
    try:
        if allowed_values:
            prompt = f"""
            Op basis van de volgende productomschrijving:
            {post_content}

            Kies een waarde voor het attribuut '{attribute_name}' uit de volgende lijst:
            {', '.join(allowed_values)}

            Geef één waarde, exact zoals in de lijst.
            """
        else:
            prompt = f"""
            Op basis van de volgende productomschrijving:
            {post_content}

            Geef een duidelijke, korte waarde voor het attribuut '{attribute_name}'.
            Ik verwacht alleen de waarde als output en geen tekst ervoor.
            Vermijd eenheden zoals 'cm', 'sets' of andere beschrijvingen. Gebruik alleen relevante kernwaarden (zoals een getal, kleur of materiaal).
            """

        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=50,
            temperature=0.7
        )

        value = response.choices[0].message.content.strip()

        # Controleer of de waarde geldig is (indien allowed_values gedefinieerd is)
        if allowed_values and value not in allowed_values:
            print(f"Waarde '{value}' is ongeldig voor '{attribute_name}'.")
            return allowed_values[0] if allowed_values else ""
        
        return value

    except Exception as e:
        print(f"Fout bij het genereren van waarde voor {attribute_name}: {e}")
        return f"Fout bij {attribute_name}"
    

def map_category(excel_category):
    return category_mapping.get(excel_category, None)

def fill_attribute_values(file_path, products_data):
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Fout bij het lezen van het Excel-bestand {file_path}: {e}")
        return

    if 'tax:product_cat' not in df.columns:
        print(f"De kolom 'tax:product_cat' ontbreekt in {file_path}.")
        return

    all_attributes = set()
    for category_name in df['tax:product_cat'].unique():
        json_category = map_category(category_name)
        if not json_category:
            print(f"Geen mapping gevonden voor categorie: {category_name}")
            continue

        required_attributes = get_required_attributes(json_category, products_data)
        all_attributes.update((attr['id'], attr['lovId']) for attr in required_attributes)

    for attr_id, _ in all_attributes:
        if attr_id not in df.columns:
            df[attr_id] = ''

    for index, row in df.iterrows():
        excel_category = row['tax:product_cat']
        json_category = map_category(excel_category)

        if not json_category:
            continue

        required_attributes = get_required_attributes(json_category, products_data)
        post_content = row.get('post_content', '')

        for attribute in required_attributes:
            attr_id = attribute['id']
            lov_id = attribute.get('lovId')

            if pd.isna(row[attr_id]) or row[attr_id] == "":
                allowed_values = get_attribute_values(lov_id, products_data) if lov_id else None
                generated_value = generate_attribute_value(post_content, attr_id, allowed_values)
                df.at[index, attr_id] = generated_value

    output_file = os.path.join(output_folder, os.path.basename(file_path))
    try:
        df.to_excel(output_file, index=False)
        print(f"Bestand verwerkt en opgeslagen als {output_file}")
    except Exception as e:
        print(f"Fout bij het opslaan van het bestand {file_path}: {e}")

# Laad het JSON-bestand
products_data = load_json_file(json_path)

# Verwerk alle bestanden in de inputmap
for file_name in os.listdir(input_folder):
    if file_name.endswith('.xlsx'):
        print(f"Verwerken: {file_name}")
        file_path = os.path.join(input_folder, file_name)
        fill_attribute_values(file_path, products_data)
