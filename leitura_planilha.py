import pandas as pd
import re
from dotenv import load_dotenv

load_dotenv()

URL_CSV = 'https://docs.google.com/spreadsheets/d/1Lx7KaQPF0oNFpbwk1dyCIVlsV-2NUGZtqc-Gn_TQfJs/export?format=csv&gid=98578301'

def normalize_text(text):
    text = str(text).lower()
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def carregar_inscricoes_pendentes():
    df = pd.read_csv(URL_CSV)
    df = df.drop(index=0).reset_index(drop=True)
    df['Inscrição'] = df['Inscrição'].apply(normalize_text)

    colunas_interesse = [
        'Numero identificação EAD',
        'Endereço de e-mail',
        'Quais cursos gostaria de realizar sua inscrição?',
        'Qual o seu nome completo?'
    ]

    df_pendentes = df[df['Inscrição'] == 'nada']
    df_pendentes = df_pendentes.dropna(subset=colunas_interesse, how='all')
    
    return df_pendentes
