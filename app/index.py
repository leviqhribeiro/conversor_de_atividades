import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pdfplumber
import pandas as pd
import os
from io import BytesIO

# Configuração da página do Streamlit
st.set_page_config(layout="wide", page_title="Auditar Engenharia")

# Interface do usuário na barra lateral
with st.sidebar:
    st.header("CONVERSOR DE ATIVIDADES")
    arquivo_pedido = st.file_uploader(
        label="Selecione o Arquivo PDF ou EXCEL das atividades:", 
        type=['pdf', 'xlsx']
    )

url = ""

conn = st.connection("gsheets", type=GSheetsConnection)

# Função para extrair as datas de início e término do PDF
def extrair_datas(caminho_pdf):
    atividades, data_inicio, data_termino = [], [], []
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            try:
                tabela = pagina.extract_table()
                if tabela:
                    df = pd.DataFrame(tabela[1:], columns=tabela[0])
                    if 'Atividade' in df.columns and 'Data de Inicio' in df.columns and 'Data Termino' in df.columns:
                        atividades.extend(df['Atividade'].dropna().tolist())
                        data_inicio.extend(df['Data de Inicio'].dropna().tolist())
                        data_termino.extend(df['Data Termino'].dropna().tolist())
            except Exception as e:
                print(f"Erro ao processar a página: {e}")
    return atividades, data_inicio, data_termino

# Função para calcular os dias de atividade e repetir o nome da atividade
def calcular_dias_atividade(df):
    df['Data de Inicio'] = pd.to_datetime(df['Data de Inicio'], format='%d/%m/%Y')
    df['Data Termino'] = pd.to_datetime(df['Data Termino'], format='%d/%m/%Y')
    df['Dias de Atividade'] = (df['Data Termino'] - df['Data de Inicio']).dt.days + 1
    df_repetida = pd.DataFrame()
    for _, row in df.iterrows():
        df_temp = pd.DataFrame({
            'Atividade': [row['Atividade']] * row['Dias de Atividade'],
            'Data de Inicio': [row['Data de Inicio']] * row['Dias de Atividade'],
            'Data Termino': [row['Data Termino']] * row['Dias de Atividade']
        })
        df_repetida = pd.concat([df_repetida, df_temp], ignore_index=True)
    return df_repetida

# Função para converter DataFrame em um arquivo Excel e retornar como um buffer em memória
def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Atividades')
    writer.save()
    processed_data = output.getvalue()
    return processed_data

# Verificação e processamento do arquivo carregado
if arquivo_pedido is not None:
    if arquivo_pedido.type == "application/pdf":
        with st.spinner('Processando arquivo PDF...'):
            with open("temp.pdf", "wb") as f:
                f.write(arquivo_pedido.read())
            
            atividades, data_inicio, data_termino = extrair_datas("temp.pdf")
            os.remove("temp.pdf")
            
            if atividades and data_inicio and data_termino:
                dias_de_atividade = pd.DataFrame({
                    'Atividade': atividades,
                    'Data de Inicio': data_inicio,
                    'Data Termino': data_termino
                })
                dias_de_atividade = calcular_dias_atividade(dias_de_atividade)
                st.dataframe(dias_de_atividade)
                
                # Adicionar botão de download
                excel_data = to_excel(dias_de_atividade)
                st.download_button(
                    label="Download do arquivo final em Excel",
                    data=excel_data,
                    file_name='dias_de_atividade.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            else:
                st.error("Atividades, datas de início e/ou término não encontradas no arquivo PDF.")
    
    elif arquivo_pedido.type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
        with st.spinner('Processando arquivo Excel...'):
            df = pd.read_excel(arquivo_pedido)
            if 'Atividade' in df.columns and 'Data de Inicio' in df.columns and 'Data Termino' in df.columns:
                df = calcular_dias_atividade(df)
                st.dataframe(df)
                
                # Adicionar botão de download
                excel_data = to_excel(df)
                st.download_button(
                    label="Download do arquivo final em Excel",
                    data=excel_data,
                    file_name='dias_de_atividade.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            else:
                st.error("As colunas 'Atividade', 'Data de Inicio' e/ou 'Data Termino' não foram encontradas no arquivo Excel.")
    
    else:
        st.error("Tipo de arquivo não suportado. Por favor, carregue um arquivo PDF ou Excel.")
