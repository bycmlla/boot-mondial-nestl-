import pandas as pd
import os

def tratar_relatorio(temp_dir, caminho_destino):
    arquivos = [f for f in os.listdir(temp_dir) if f.endswith(".xls")]
    if not arquivos:
        raise FileNotFoundError("Nenhum arquivo .xls encontrado no diretório temporário!")

    caminho_xls = os.path.join(temp_dir, arquivos[0])

    tabelas = pd.read_html(caminho_xls, decimal=",", thousands=".")
    df = tabelas[0]

    print("Prévia dos dados:")
    print(df.head())

    df.to_excel(caminho_destino, index=False)

    print(f"✅ Dados salvos em: {caminho_destino}")
