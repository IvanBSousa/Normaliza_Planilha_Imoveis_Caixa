import pandas as pd
import os
import re
from openpyxl.utils import get_column_letter
from datetime import date

diretorio = os.path.dirname("C:\\Users\\ivanb\\Downloads\\")

def trata_csv_imoveis(nome_arquivo):
    # Lê o arquivo CSV
    caminho_arquivo = os.path.join(diretorio, nome_arquivo)
    print(f"Lendo arquivo CSV: {caminho_arquivo}")
    encodings = ['cp1252', 'latin-1', 'utf-8']
    df = None
    last_error = None

    for encoding in encodings:
        try:
            df = pd.read_csv(
                caminho_arquivo,
                encoding=encoding,
                sep=';',
                skiprows=2,
                skip_blank_lines=True,
                skipinitialspace=True,
                engine='python',
            )
            df.columns = df.columns.str.strip()
            print(f"Lido com encoding={encoding} e sep=';' (skiprows=2)")
            break
        except UnicodeDecodeError as e:
            last_error = e
        except pd.errors.ParserError as e:
            last_error = e

    if df is None:
        raise last_error

    # Converte Preço e Valor de avaliação para numérico
    valor_cols = ['Preço', 'Valor de avaliação']
    for col in valor_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False).str.strip()
            df[col] = pd.to_numeric(df[col], errors='coerce')

    # Normaliza a coluna Endereço para manter apenas rua e número
    if 'Endereço' in df.columns:
        cleanup_pattern = re.compile(
            r'(?i)\b(?:apt(?:\.|o)?|apto|apartamento|casa|lote|lt(?:\.|)?|quadra|qd(?:\.|)?|q\.d\.?|bloco|bl(?:\.|)?|conjunto|cjto|campus|andar|sala|box|frente|fundos|cond(?:\.|om[ií]nio)?)\b.*$',
            re.IGNORECASE,
        )

        def only_street_number(value):
            value = str(value).strip()
            value = re.sub(r'\s*,\s*', ' ', value)
            value = re.sub(r'\s+', ' ', value).strip()
            value = re.sub(r'\bN\.?\s*', '', value, flags=re.IGNORECASE)
            value = re.sub(r'\bS\/N?\s*', '', value, flags=re.IGNORECASE)
            value = re.sub(r'\bSN?\s', '', value, flags=re.IGNORECASE)
            street_part = cleanup_pattern.sub('', value).strip()
            street_part = re.sub(r'[\s,;:-]+$', '', street_part)
            return street_part

        df['Endereço'] = df['Endereço'].apply(only_street_number)

    # Cria coluna Endereços unindo Endereço, Bairro, Cidade e UF
    endereco_cols = ['Endereço', 'Bairro', 'Cidade', 'UF']
    missing_cols = [col for col in endereco_cols if col not in df.columns]
    if missing_cols:
        raise KeyError(f"Colunas faltando para montar 'Endereços': {missing_cols}")

    df['Endereços'] = (
        df['Endereço'].astype(str).str.strip()
        + ' ' + df['Bairro'].astype(str).str.strip()
        + ' ' + df['Cidade'].astype(str).str.strip()
        + ' ' + df['UF'].astype(str).str.strip()
    )

    # Remove a coluna de número de imóvel
    drop_cols = [col for col in ['N° do imóvel', 'Nº do imóvel', 'No do imóvel'] if col in df.columns]
    if drop_cols:
        df.drop(columns=drop_cols, inplace=True)

    # Reordena as colunas para deixar Endereços primeiro
    cols = df.columns.tolist()
    cols.remove('Endereços')
    cols = ['Endereços'] + cols
    df = df[cols]

    df.drop(columns=['Endereço', 'Bairro', 'Cidade', 'UF'], inplace=True)

    # Salva em XLSX com formatação de moeda nas colunas de valor
    nome_saida = os.path.splitext(nome_arquivo)[0] + '_' + str(date.today()) + '.xlsx'
    caminho_saida = os.path.join(diretorio, nome_saida)

    with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Imóveis')
        workbook = writer.book
        worksheet = writer.sheets['Imóveis']

        money_format = 'R$ #,##0.00'
        for col_idx, col_name in enumerate(df.columns, start=1):
            if col_name in ['Preço', 'Valor de avaliação']:
                column_letter = get_column_letter(col_idx)
                for row in range(2, len(df) + 2):
                    worksheet[f'{column_letter}{row}'].number_format = money_format

    print(f"Salvo XLSX em: {caminho_saida}")

    return df

trata_csv_imoveis("Lista_imoveis_MG.csv")