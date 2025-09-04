import shutil
import zipfile
import logging
import pandas as pd
import os
import glob
from formatacao_excel import format_excel_sheet

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def treatment_csv(download_dir):
    try:
        zip_files = glob.glob(os.path.join(download_dir, "*.zip"))
        if not zip_files:
            logging.warning("Nenhum arquivo ZIP encontrado no diretório: %s", download_dir)
            return None

        zip_file_path = zip_files[0]
        logging.info(f"Descompactando arquivo: {zip_file_path}")
        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            zip_ref.extractall(download_dir)

        csv_files = glob.glob(os.path.join(download_dir, "*.csv"))
        if not csv_files:
            logging.warning("Nenhum arquivo CSV encontrado após descompactar o ZIP.")
            return None

        csv_file_path = csv_files[0]
        logging.info(f"Lendo arquivo CSV: {csv_file_path}")

        try:
            df = pd.read_csv(csv_file_path, encoding='cp1252', on_bad_lines='warn', sep=';')
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(csv_file_path, encoding='utf-8', on_bad_lines='warn', sep=';')
                logging.info("CSV lido com encoding utf-8.")
            except Exception as e:
                logging.error(f"Erro ao ler o CSV com ambas as codificações (cp1252, utf-8): {e}")
                return None
        except Exception as e:
            logging.error(f"Erro inesperado ao ler o CSV: {e}")
            return None

        df = df[df["FATURA"].isin(["", "-"]) | df["FATURA"].isna()]
        df = df[~df["TNOT_DESCRICAO_STATUS"].astype(str).str.contains("CANCELADO", na=False)]

        df["TCON_DATA_EMISSAO"] = pd.to_datetime(df["TCON_DATA_EMISSAO"], errors='coerce', dayfirst=True)
        df.dropna(subset=["TCON_DATA_EMISSAO"], inplace=True)
        df['Mes'] = df['TCON_DATA_EMISSAO'].dt.month.astype(int)
        df['Dia'] = df['TCON_DATA_EMISSAO'].dt.day.astype(int)
        logging.info("Colunas 'Mes' e 'Dia' criadas.")

        df["TCON_VALOR_LIQUIDO"] = df["TCON_VALOR_LIQUIDO"].astype(str).str.replace('.', '', regex=False).str.replace(
            ',', '.')
        df["TCON_VALOR_LIQUIDO"] = pd.to_numeric(df["TCON_VALOR_LIQUIDO"], errors='coerce')
        df.dropna(subset=["TCON_VALOR_LIQUIDO"], inplace=True)
        logging.info("Coluna 'TCON_VALOR_LIQUIDO' convertida para numérico.")

        logging.info("Criando resumo (pivot table) com Pandas (agrupado por Mês)...")
        try:
            pivot_df = pd.pivot_table(
                df,
                index=['Mes'],
                values='TCON_VALOR_LIQUIDO',
                aggfunc='sum',
                fill_value=0
            )
            pivot_df.rename(columns={'TCON_VALOR_LIQUIDO': 'Soma de Valor Líquido'}, inplace=True)

            month_map = {
                1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
                5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
                9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
            }
            pivot_df.index = pivot_df.index.map(month_map)
        except KeyError as ke:
            logging.error(
                f"Erro ao criar pivot table: Coluna não encontrada - {ke}. Colunas disponíveis: {df.columns.tolist()}")
            return None
        except Exception as e:
            logging.error(f"Erro inesperado ao criar pivot table pandas: {e}", exc_info=True)
            return None

        xlsx_file_path = os.path.join(download_dir, "MONDIAL SEM FATURAS.xlsx")
        sheet_name_data = 'Dados'
        sheet_name_pivot = 'ResumoMensal'

        logging.info(f"Salvando dados e resumo mensal em: {xlsx_file_path}")
        try:
            with pd.ExcelWriter(xlsx_file_path, engine='openpyxl',
                                datetime_format='DD/MM/YYYY',
                                date_format='DD/MM/YYYY') as writer:
                df.to_excel(writer, index=False, sheet_name=sheet_name_data)
                logging.info(f"DataFrame original salvo na planilha '{sheet_name_data}'.")

                pivot_df.to_excel(writer, sheet_name=sheet_name_pivot)
                logging.info(f"Resumo Mensal salvo na planilha '{sheet_name_pivot}'.")

                workbook = writer.book
                ws_data = workbook[sheet_name_data]
                format_excel_sheet(ws_data, df.columns.tolist())

                ws_pivot = workbook[sheet_name_pivot]
                pivot_cols_for_formatting = pivot_df.index.names + pivot_df.columns.tolist()
                format_excel_sheet(ws_pivot, pivot_cols_for_formatting)
            try:
                download_path = os.path.join(os.path.expanduser("~"), "Downloads", "MONDIAL SEM FATURAS.xlsx")
                shutil.copy(xlsx_file_path, download_path)
                logging.info(f"Cópia do arquivo salva em: {download_path}")
            except Exception as e:
                logging.error(f"Erro ao copiar o arquivo para Downloads: {e}")

            logging.info(f"Arquivo XLSX final '{xlsx_file_path}' criado com ambas as planilhas e formatação essencial.")
            return xlsx_file_path

        except Exception as e:
            logging.error(f"Erro ao escrever ou formatar o arquivo Excel: {e}", exc_info=True)
            return None

    except FileNotFoundError as fnf_err:
        logging.error(f"Erro de arquivo não encontrado durante o tratamento: {fnf_err}")
        return None
    except zipfile.BadZipFile:
        logging.error(f"Erro ao descompactar: O arquivo {zip_file_path} não é um ZIP válido ou está corrompido.")
        return None
    except KeyError as ke:
        logging.error(
            f"Erro de chave/coluna não encontrada no DataFrame: {ke}. Verifique os nomes das colunas no arquivo CSV/DataFrame.")
        return None
    except Exception as e:
        logging.error(f"Erro geral inesperado na função treatment_csv: {e}", exc_info=True)
        return None
