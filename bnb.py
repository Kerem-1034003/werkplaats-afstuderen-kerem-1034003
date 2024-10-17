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

