import requests
import pandas as pd
from datetime import datetime
from google.oauth2.service_account import Credentials
import gspread

URL_API = "http://10.0.0.111:30800/v1/md?data=20240310"
ABA_NOME = "acompanhamento"
CAMINHO_CREDENCIAIS = "credenciais.json"
KEY = "1cvVZ0CporgR_-OJGsiIUk-SE3GF5g3CQ0KJYDP09AXM"

DATA_LIMITE = datetime.strptime("31/07/2025", "%d/%m/%Y")

def atualizar_google_sheets():
    response = requests.get(URL_API)
    response.raise_for_status()

    dados = response.json()

    print("=== 5 primeiras linhas da API ===")
    for i, item in enumerate(dados[:5]):
        print(item)
    print("=== Fim ===")

    df = pd.DataFrame(dados)

    if 'dataemissao' in df.columns:
        # Ajusta para datetime
        df['dataemissao'] = pd.to_datetime(df['dataemissao'], errors='coerce')

        df = df[df['dataemissao'] > DATA_LIMITE]

        df['dataemissao'] = df['dataemissao'].dt.strftime("%d/%m/%Y")

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    credentials = Credentials.from_service_account_file(CAMINHO_CREDENCIAIS, scopes=scopes)
    client = gspread.authorize(credentials)

    spreadsheet = client.open_by_key(KEY)
    sheet = spreadsheet.worksheet(ABA_NOME)

    sheet.clear()
    sheet.append_row(df.columns.tolist())
    if not df.empty:
        sheet.append_rows(df.values.tolist())

    print(f"✅ Dados atualizados na aba '{ABA_NOME}' com {len(df)} registros.")

if __name__ == "__main__":
    atualizar_google_sheets()
