import os
from dotenv import load_dotenv
import pandas as pd
import openai

load_dotenv()

openai_api_key = os.getenv('OPENAI_API_KEY')

openai.api_key = openai_api_key

df = pd.read_excel('excel/Autoinkoop.xlsx')