import streamlit as st
import pdfplumber
import pandas as pd
import os
from io import BytesIO
from openpyxl import Workbook
import numpy as np

# Configuração da página
st.set_page_config(layout="wide", page_title="Auditar Engenharia")

# Interface do usuário na barra lateral
with st.sidebar:
    st.header("CONVERSOR DE ATIVIDADES")
    arquivo_atividades = st.file_uploader(
        label="Selecione o arquivo PDF ou EXCEL das Atividades:",
        type=("pdf", "xlsx")
    )

# Definindo possíveis nomes das colunas
possiveis_nomes_colunas = {
    "Atividade": ["ITEM", "Nome da Tarefa", "Nome da Atividade", "Atividade", "Tarefa"],
    "Data Inicio": ["Início", "INÍCIO", "Data de Inicio", "Data Inicial", "Data a ser Iniciada"],
    "Data Termino": ["Término", "TÉRMINO", "Data de Termino", "Data Final", "Data a ser Concluida", "Data Conclusao"]
}

# Função para renomear colunas do DataFrame
def renomear_colunas(df, possiveis_nomes_colunas):
    nova_colunas = {}
    for nome_padrao, nomes_alternativos in possiveis_nomes_colunas.items():
        for nome in nomes_alternativos:
            if nome in df.columns:
                nova_colunas[nome] = nome_padrao
                break
    return df.rename(columns=nova_colunas)

# Função para extrair as datas de início e término do PDF
def extrair_datas(caminho_pdf):
    atividades, data_inicio, data_termino = [], [], []
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            try:
                tabela = pagina.extract_table()
                if tabela:
                    df = pd.DataFrame(tabela[1:], columns=tabela[0])
                    df = renomear_colunas(df, possiveis_nomes_colunas)
                    if 'Atividade' in df.columns and 'Data Inicio' in df.columns and 'Data Termino' in df.columns:
                        atividades.extend(df['Atividade'].dropna().tolist())
                        data_inicio.extend(df['Data Inicio'].dropna().tolist())
                        data_termino.extend(df['Data Termino'].dropna().tolist())
            except Exception as e:
                print(f"Erro ao processar a página: {e}")
    return atividades, data_inicio, data_termino

# Função para calcular os dias de atividade e repetir o nome da atividade
def calcular_dias_atividade(df):
    df['Data Inicio'] = pd.to_datetime(df['Data Inicio'], errors='coerce', dayfirst=True)
    df['Data Termino'] = pd.to_datetime(df['Data Termino'], errors='coerce', dayfirst=True)
    df['Dias de Atividade'] = (df['Data Termino'] - df['Data Inicio']).dt.days + 1

    df_final = pd.DataFrame()
    for _, row in df.iterrows():
        df_temp = pd.DataFrame({
            'Atividade': [row['Atividade']] * row['Dias de Atividade'],
            'Data Inicio': [row['Data Inicio']] * row['Dias de Atividade'],
            'Data Termino': [row['Data Termino']] * row['Dias de Atividade'],
            'Data de Execução': [row['Data Inicio'] + pd.Timedelta(days=i) for i in range(row['Dias de Atividade'])],
        })
        df_final = pd.concat([df_final, df_temp], ignore_index=True)

    # Formatando as colunas de data para 'dd/mm/yyyy' e removendo as horas
    df_final['Data Inicio'] = df_final['Data Inicio'].dt.strftime('%d/%m/%Y')
    df_final['Data Termino'] = df_final['Data Termino'].dt.strftime('%d/%m/%Y')
    df_final['Data de Execução'] = df_final['Data de Execução'].dt.strftime('%d/%m/%Y')

    return df_final

# Função para converter DataFrame em um arquivo Excel
def to_excel(df, nome_da_planilha_principal= 'Atividades'):    
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name= nome_da_planilha_principal)
    writer.save()
    processed_data = output.getvalue()
    return processed_data

# Verificação e processamento do arquivo carregado
if arquivo_atividades is not None:
    if arquivo_atividades.type == "application/pdf":
        with st.spinner('Processando arquivo PDF...'):
            with open("temp.pdf", "wb") as arquivo_temporario:
                arquivo_temporario.write(arquivo_atividades.read())

            atividades, data_inicio, data_termino = extrair_datas("temp.pdf")
            os.remove("temp.pdf")

            if atividades and data_inicio and data_termino:
                dias_de_atividade = pd.DataFrame({
                    'Atividade': atividades,
                    'Data Inicio': data_inicio,
                    'Data Termino': data_termino
                })
                dias_de_atividade = calcular_dias_atividade(dias_de_atividade)
                st.dataframe(dias_de_atividade)

                # Convertendo a coluna 'Data de Execução' para datetime
                dias_de_atividade['Data de Execução'] = pd.to_datetime(dias_de_atividade['Data de Execução'], format='%d/%m/%Y', errors='coerce')

                # Lista de datas únicas
                datas_unicas = dias_de_atividade['Data de Execução'].unique()
                data_selecionada = st.selectbox('Selecione a data para download:', options=sorted(datas_unicas))

                if data_selecionada is not None:
                    atividades_data_selecionada = dias_de_atividade[dias_de_atividade['Data de Execução'] == data_selecionada]

                    # Adicionar botão de download
                    excel_data = to_excel(atividades_data_selecionada)
                    st.download_button(
                        label=f"Download do arquivo Excel para {pd.to_datetime(data_selecionada).strftime('%d/%m/%Y')}",
                        data=excel_data,
                        file_name=f'atividades_{pd.to_datetime(data_selecionada).strftime("%d-%m-%Y")}.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
            else:
                st.error("Atividade, Datas de início e/ou término não encontradas no arquivo PDF.")

    elif arquivo_atividades.type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
        with st.spinner('Processando arquivo Excel...'):
            df = pd.read_excel(arquivo_atividades)
            df = renomear_colunas(df, possiveis_nomes_colunas)
            if 'Atividade' in df.columns and 'Data Inicio' in df.columns and 'Data Termino' in df.columns:
                df = calcular_dias_atividade(df)
                st.dataframe(df)

                # Convertendo a coluna 'Data de Execução' para datetime
                df['Data de Execução'] = pd.to_datetime(df['Data de Execução'], format='%d/%m/%Y', errors='coerce', dayfirst=True)
                df['Data de Execução'] = df['Data de Execução'].dt.strftime('%d/%m/%Y')
                # Lista de datas únicas
                datas_unicas = df['Data de Execução'].unique()
                data_selecionada = st.selectbox('Selecione a data para download:', options=sorted(datas_unicas))

                if data_selecionada is not None:
                    atividades_data_selecionada = df[df['Data de Execução'] == data_selecionada]

                    # Adicionar botão de download
                    excel_data = to_excel(atividades_data_selecionada)
                    st.download_button(
                        label=f"Download do arquivo Excel para {data_selecionada}",
                        data=excel_data,
                        file_name=f'atividades_{data_selecionada.replace("/", "-")}.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
            else:
                st.error("As colunas necessárias (Atividade, Data Inicio, Data Termino) não foram encontradas no arquivo Excel.")
else:
    st.warning("Escolha um arquivo para começar.")
