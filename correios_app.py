import streamlit as st
import pandas as pd
import pdfplumber
import re
import requests
import time
import io
import numpy as np

# Configuração da página
st.set_page_config(page_title="Ferramentas de Logística", layout="wide")

# Título Principal
st.title("📦 Ferramentas de Processamento - Logística")

# Criando abas para separar os dois sistemas
aba1, aba2 = st.tabs(["📄 Extrator de Rastreios (PDF)", "📍 Processador de Planilhas (CEP)"])

# ==========================================
# ABA 1: EXTRATOR DE PDFs DOS CORREIOS
# ==========================================
with aba1:
    st.header("Extrair Códigos de Rastreio dos Correios")
    st.write("Faça o upload dos PDFs gerados pelos Correios para gerar a planilha com Nomes e Rastreiamentos.")
    
    arquivos_pdf = st.file_uploader("Selecione os PDFs", type=["pdf"], accept_multiple_files=True)
    
    if arquivos_pdf and st.button("Processar PDFs"):
        dados_extraidos = []
        padrao_rastreio = re.compile(r'(\S{2}\d{9}BR)', re.IGNORECASE)
        
        with st.spinner("Lendo PDFs... Isso pode levar alguns segundos."):
            for arquivo in arquivos_pdf:
                try:
                    with pdfplumber.open(arquivo) as pdf:
                        for pagina in pdf.pages:
                            texto = pagina.extract_text()
                            if not texto: continue
                            
                            linhas = texto.split('\n')
                            for i, linha in enumerate(linhas):
                                match = padrao_rastreio.search(linha)
                                if match:
                                    codigo_rastreio = match.group(1).upper()
                                    resto_linha = linha.replace(match.group(0), '').strip()
                                    nome_pintor = re.split(r'\s\d{4,}', resto_linha)[0].strip()
                                    
                                    if not re.search(r'[A-Za-z]', nome_pintor):
                                        for offset in [1, 2, 3]:
                                            if (i + offset) < len(linhas):
                                                linha_abaixo = linhas[i + offset].strip()
                                                if re.search(r'[A-Za-z]', linha_abaixo) and not padrao_rastreio.search(linha_abaixo):
                                                    nome_pintor = re.split(r'\s\d{4,}', linha_abaixo)[0].strip()
                                                    break
                                    
                                    nome_pintor = re.sub(r'\b(PAC|SEDEX).*', '', nome_pintor, flags=re.IGNORECASE).strip()
                                    
                                    if nome_pintor and "Código do objeto" not in nome_pintor:
                                        dados_extraidos.append({
                                            'Nome do Pintor': nome_pintor,
                                            'Código de Rastreio': codigo_rastreio
                                        })
                except Exception as e:
                    st.error(f"Erro ao processar {arquivo.name}: {e}")

        if dados_extraidos:
            df_rastreios = pd.DataFrame(dados_extraidos).drop_duplicates()
            st.success(f"Sucesso! {len(df_rastreios)} registros extraídos.")
            st.dataframe(df_rastreios.head()) # Mostra uma prévia na tela
            
            # Prepara o Excel para download em memória
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer: