import requests
import pandas as pd
from google.oauth2.service_account import Credentials
import gspread

URL_API = "http://10.0.0.3:30800/v1/md?data=20250801"
ABA_NOME = "Acompanhamento"
CAMINHO_CREDENCIAIS = "credenciais.json"
KEY = "1hTgeKcoVtBrDbIh86eOtL8x0fO0gecjJd1RBnGHyz6c"


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
        df['dataemissao'] = pd.to_datetime(df['dataemissao'], errors='coerce')
        df['dataemissao'] = df['dataemissao'].dt.strftime("%d/%m/%Y")

    tomadores = [
        "NESTLE BRASIL LTDA",
        "NESTLE BRASIL LTDA.",
        "Nestle Brasil Ltda.",
        "NESTLE DO BRASIL LTDA",
        "Nestle Nordeste Alimentos e Bebidas",
        "NESTLE NORDESTE ALIMENTOS E BEBIDAS LTDA."
    ]
    df = df[df['tomador'].isin(tomadores)]

    status_validos = ["BASE", "TRANSITO"]
    df = df[df['statusentrega'].isin(status_validos)]

    df['CONCAT'] = df['dataemissao'].astype(str) + " " + df['cnpjdestino'].astype(str)

    colunas_para_remover = [
        "tipodoc", "razaosocial", "cnpj", "valorpgtone",
        "valortotal", "valorliquido", "saldo", "valornfe",
        "peso", "volume", "dia", "cfop", "basecalculo", "icms", "nomefantasia", "inscricaoestadual", "valorpgto",
        "fatura", "dataoco", "horaoco", "statusfatura", "mdfe", "manifesto", "cpfmotorista", "reboque1", "reboque2",
        "chavecte", "statusmn", "situacao", "aliquota"
    ]
    df = df.drop(columns=[col for col in colunas_para_remover if col in df.columns])

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    credentials = Credentials.from_service_account_file(CAMINHO_CREDENCIAIS, scopes=scopes)
    client = gspread.authorize(credentials)

    spreadsheet = client.open_by_key(KEY)
    sheet = spreadsheet.worksheet(ABA_NOME)

    # Limpa e atualiza a aba
    sheet.clear()
    sheet.append_row(df.columns.tolist())
    if not df.empty:
        sheet.append_rows(df.values.tolist())

    print(f"✅ Dados atualizados na aba '{ABA_NOME}' com {len(df)} registros.")


if __name__ == "__main__":
    atualizar_google_sheets()
