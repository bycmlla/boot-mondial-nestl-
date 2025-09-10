import os
import xml.etree.ElementTree as ET
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

PASTA_XML = r"\\srv004-jtd\Analitics\ACOMPANHAMENTO DE CARGAS\xml"
ABA_NOME = "XML"
CAMINHO_CREDENCIAIS = "credenciais.json"
KEY = "1hTgeKcoVtBrDbIh86eOtL8x0fO0gecjJd1RBnGHyz6c"

scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file(CAMINHO_CREDENCIAIS, scopes=scopes)
gc = gspread.authorize(creds)
sh = gc.open_by_key(KEY)

try:
    worksheet = sh.worksheet(ABA_NOME)
except gspread.exceptions.WorksheetNotFound:
    worksheet = sh.add_worksheet(title=ABA_NOME, rows="1000", cols="30")

if worksheet.acell('A1').value is None:
    worksheet.append_row([
        "TIPO NF", "SERIE", "NUMERO NF", "DATA EMISSÃO", "DATA SAIDA",
        "CNPJ REMETENTE", "NFe.infNFe.emit.xNome", "REMETENTE", "CIDADE REMETENTE", "UF REMETENTE",
        "NFe.infNFe.emit.IE", "NFe.infNFe.emit.IM", "NFe.infNFe.emit.CRT",
        "CNPJ DESTINO", "DESTINO", "BAIRRO DESTINO", "CIDADE DESTINO", "UF DESTINO",
        "CÓDIGO PRODUTO", "DESCRIÇÃO PRODUTO", "QUANT", "VALOR UNIT", "VALOR TOTAL", "CAMPO OBSERVAÇÃO", "CONCAT"
    ])

try:
    aba_controle = sh.worksheet("Controle")
except gspread.exceptions.WorksheetNotFound:
    aba_controle = sh.add_worksheet(title="Controle", rows="1000", cols="2")

if aba_controle.row_count == 0 or aba_controle.get_all_values() == []:
    aba_controle.append_row(["Arquivos Processados"])

arquivos_processados = aba_controle.col_values(1)[1:]

dados_para_inserir = []
novos_processados = []

inicio = datetime.now()
print(f"[{inicio}] Iniciando processamento de XMLs...")


def formatar_data(data_str):
    if not data_str:
        return ""
    try:
        dt = datetime.fromisoformat(data_str.replace("Z", "+00:00"))
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return data_str.split("T")[0]


for arquivo_nome in os.listdir(PASTA_XML):
    if not arquivo_nome.lower().endswith(".xml"):
        continue
    if arquivo_nome in arquivos_processados:
        continue

    caminho_arquivo = os.path.join(PASTA_XML, arquivo_nome)
    tree = ET.parse(caminho_arquivo)
    root = tree.getroot()
    ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

    nfe = root.find('nfe:NFe', ns)
    if nfe is None:
        continue
    infNFe = nfe.find('nfe:infNFe', ns)
    ide = infNFe.find('nfe:ide', ns)
    emit = infNFe.find('nfe:emit', ns)
    endEmit = emit.find('nfe:enderEmit', ns)
    dest = infNFe.find('nfe:dest', ns)
    endDest = dest.find('nfe:enderDest', ns)
    obs = infNFe.find('nfe:infAdic', ns)

    natOp = ide.findtext('nfe:natOp', default='', namespaces=ns)
    if natOp != "TRANSF.DE MAT.USO/CONSUMO":
        continue

    tipoNF = ide.findtext('nfe:tpNF', default='', namespaces=ns)
    serie = ide.findtext('nfe:serie', default='', namespaces=ns)
    numeroNF = ide.findtext('nfe:nNF', default='', namespaces=ns)
    dataEmissao = formatar_data(ide.findtext('nfe:dhEmi', default='', namespaces=ns))
    dataSaida = formatar_data(ide.findtext('nfe:dhSaiEnt', default='', namespaces=ns))

    cnpjEmit = emit.findtext('nfe:CNPJ', default='', namespaces=ns)
    nomeEmit = emit.findtext('nfe:xNome', default='', namespaces=ns)
    cidadeEmit = endEmit.findtext('nfe:xMun', default='', namespaces=ns)
    ufEmit = endEmit.findtext('nfe:UF', default='', namespaces=ns)
    ieEmit = emit.findtext('nfe:IE', default='', namespaces=ns)
    imEmit = emit.findtext('nfe:IM', default='', namespaces=ns)
    crtEmit = emit.findtext('nfe:CRT', default='', namespaces=ns)

    cnpjDest = dest.findtext('nfe:CNPJ', default='', namespaces=ns) or dest.findtext('nfe:CPF', default='', namespaces=ns)
    nomeDest = dest.findtext('nfe:xNome', default='', namespaces=ns)
    bairroDest = endDest.findtext('nfe:xBairro', default='', namespaces=ns)
    cidadeDest = endDest.findtext('nfe:xMun', default='', namespaces=ns)
    ufDest = endDest.findtext('nfe:UF', default='', namespaces=ns)

    campoObs = obs.findtext('nfe:infCpl', default='', namespaces=ns) if obs is not None else ''

    CONCAT = f"{dataEmissao} {cnpjDest}"

    for det in infNFe.findall('nfe:det', ns):
        prod = det.find('nfe:prod', ns)
        if prod is None:
            continue
        codProd = prod.findtext('nfe:cProd', default='', namespaces=ns)
        descProd = prod.findtext('nfe:xProd', default='', namespaces=ns)
        qtd = prod.findtext('nfe:qCom', default='', namespaces=ns)
        vUnit = prod.findtext('nfe:vUnCom', default='', namespaces=ns)
        vTotal = prod.findtext('nfe:vProd', default='', namespaces=ns)

        dados_para_inserir.append([
            tipoNF, serie, numeroNF, dataEmissao, dataSaida,
            cnpjEmit, nomeEmit, nomeEmit, cidadeEmit, ufEmit,
            ieEmit, imEmit, crtEmit,
            cnpjDest, nomeDest, bairroDest, cidadeDest, ufDest,
            codProd, descProd, qtd, vUnit, vTotal, campoObs, CONCAT
        ])

    novos_processados.append([arquivo_nome])

if dados_para_inserir:
    worksheet.append_rows(dados_para_inserir, value_input_option='USER_ENTERED')

if novos_processados:
    aba_controle.append_rows(novos_processados, value_input_option='USER_ENTERED')

fim = datetime.now()
print(f"[{fim}] Processamento finalizado.")
print(f"Total de arquivos processados: {len(novos_processados)}")
print(f"Tempo total de execução: {(fim - inicio).total_seconds()} segundos")
