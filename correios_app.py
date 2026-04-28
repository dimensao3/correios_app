import streamlit as st
import pandas as pd
import pdfplumber
import re
import requests
import time
import io
import numpy as np
from datetime import datetime

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
            # =========================================================
            # ALTERAÇÃO FEITA: Removido o .drop_duplicates()
            # Agora ele pega absolutamente TUDO, sem tirar nem por.
            # =========================================================
            df_rastreios = pd.DataFrame(dados_extraidos)
            
            st.success(f"Sucesso! {len(df_rastreios)} registros extraídos (incluindo possíveis repetições do PDF).")
            st.dataframe(df_rastreios.head()) 
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_rastreios.to_excel(writer, index=False)
            
            data_hora_atual = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
            nome_arquivo_rastreios = f"Planilha_Rastreios_{data_hora_atual}.xlsx"
            
            st.download_button(
                label="📥 Baixar Planilha de Rastreios",
                data=buffer.getvalue(),
                file_name=nome_arquivo_rastreios,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nenhum dado encontrado nos PDFs selecionados.")

# ==========================================
# ABA 2: PROCESSADOR DE CEP E FORMATAÇÃO
# ==========================================
with aba2:
    st.header("Processar Planilha Base de Resgates")
    st.write("Faça o upload da planilha base. O sistema manterá a coluna 'ENDEREÇO' original entre o CEP e a Rua.")
    
    arquivo_base = st.file_uploader("Selecione a Planilha Base", type=["csv", "xlsx"])
    
    if arquivo_base and st.button("Processar Planilha"):
        
        def buscar_cep(cep):
            if not cep or len(str(cep)) != 8: return None
            headers = {'User-Agent': 'Mozilla/5.0'}
            try:
                r = requests.get(f"https://viacep.com.br/ws/{cep}/json/", headers=headers, timeout=5)
                if r.status_code == 200 and "erro" not in r.json():
                    return {'logradouro': r.json().get('logradouro', ''), 'bairro': r.json().get('bairro', ''), 'localidade': r.json().get('localidade', ''), 'uf': r.json().get('uf', '')}
            except: pass
            
            try:
                r = requests.get(f"https://brasilapi.com.br/api/cep/v1/{cep}", headers=headers, timeout=5)
                if r.status_code == 200:
                    return {'logradouro': r.json().get('street', ''), 'bairro': r.json().get('neighborhood', ''), 'localidade': r.json().get('city', ''), 'uf': r.json().get('state', '')}
            except: pass
            return None

        def limpa_numero(valor, tamanho):
            txt = str(valor).strip()
            if txt.endswith('.0'): txt = txt[:-2]
            txt = ''.join(filter(str.isdigit, txt))    
            if len(txt) > 0 and txt != '0': return txt.zfill(tamanho)
            return ""

        with st.spinner("Carregando e formatando dados..."):
            if arquivo_base.name.endswith('.csv'):
                try:
                    df = pd.read_csv(arquivo_base, sep=',', encoding='utf-8')
                except UnicodeDecodeError:
                    df = pd.read_csv(arquivo_base, sep=',', encoding='latin1')
            else:
                df = pd.read_excel(arquivo_base)

            df.columns = df.columns.str.strip()
            df = df.dropna(subset=['ID PINTOR'])
            
            df['CPF'] = df['CPF'].apply(lambda x: limpa_numero(x, 11))
            df['CEP'] = df['CEP'].apply(lambda x: limpa_numero(x, 8))
            
            if 'ENDEREÇO' not in df.columns:
                df['ENDEREÇO'] = ""

            df = df[df['CEP'] != ""]

            for col in ['PESO', 'ALTURA', 'LARGURA', 'COMPRIMENTO']:
                if col in df.columns:
                    df[col] = df[col].astype(str).replace(r'\.0$', '', regex=True).replace(r"'+", "", regex=True).replace('nan', '').fillna("")
                else:
                    df[col] = ""

            agg_funcs = {
                'CPF': 'first', 
                'PINTOR': 'first', 
                'CEP': 'first', 
                'ENDEREÇO': 'first',
                'PESO': 'max', 
                'ALTURA': 'max', 
                'LARGURA': 'max', 
                'COMPRIMENTO': 'max'
            }
            df_fixo = df.groupby('ID PINTOR').agg(agg_funcs).reset_index()

            df['Material_Idx'] = df.groupby('ID PINTOR').cumcount() + 1
            df_premio = df.pivot(index='ID PINTOR', columns='Material_Idx', values='MATERIAL').add_prefix('PREMIO_')
            df_qtd = df.pivot(index='ID PINTOR', columns='Material_Idx', values='QUANTIDADE').add_prefix('QUANTIDADE_')

            df_materiais = pd.DataFrame(index=df_premio.index)
            materiais_cols = []
            for i in range(1, df['Material_Idx'].max() + 1):
                if f'PREMIO_{i}' in df_premio.columns:
                    df_materiais[f'PREMIO_{i}'] = df_premio[f'PREMIO_{i}']
                    materiais_cols.append(f'PREMIO_{i}')
                if f'QUANTIDADE_{i}' in df_qtd.columns:
                    df_materiais[f'QUANTIDADE_{i}'] = df_qtd[f'QUANTIDADE_{i}']
                    materiais_cols.append(f'QUANTIDADE_{i}')
                df_materiais[f'VALOR_{i}'] = ""
                materiais_cols.append(f'VALOR_{i}')

            df_final = pd.merge(df_fixo, df_materiais, on='ID PINTOR', how='left')
            ceps_unicos = df_final['CEP'].replace("", np.nan).dropna().unique()
            dic_ceps = {}

        st.write(f"Consultando {len(ceps_unicos)} CEPs únicos...")
        barra_progresso = st.progress(0)
        
        for idx, cep in enumerate(ceps_unicos):
            if len(cep) == 8:
                dic_ceps[cep] = buscar_cep(cep)
                time.sleep(0.3)
            barra_progresso.progress((idx + 1) / len(ceps_unicos))

        with st.spinner("Finalizando formatação..."):
            df_final['RUA'] = df_final['CEP'].apply(lambda c: dic_ceps.get(c, {}).get('logradouro') if dic_ceps.get(c) else 'NÃO ENCONTRADO')
            df_final['BAIRRO'] = df_final['CEP'].apply(lambda c: dic_ceps.get(c, {}).get('bairro') if dic_ceps.get(c) else '')
            df_final['CIDADE'] = df_final['CEP'].apply(lambda c: dic_ceps.get(c, {}).get('localidade') if dic_ceps.get(c) else '')
            df_final['UF'] = df_final['CEP'].apply(lambda c: dic_ceps.get(c, {}).get('uf') if dic_ceps.get(c) else '')
            df_final['NUMERO_RESIDENCIA'] = ""
            df_final['COMPLEMENTO'] = ""

            ordem_base = [
                'ID PINTOR', 'CPF', 'PINTOR', 'CEP', 'ENDEREÇO', 'RUA', 
                'NUMERO_RESIDENCIA', 'COMPLEMENTO', 'BAIRRO', 'CIDADE', 
                'UF', 'PESO', 'ALTURA', 'LARGURA', 'COMPRIMENTO'
            ]
            df_final = df_final[ordem_base + materiais_cols]

            st.success("Planilha processada com sucesso!")
            st.dataframe(df_final.head())

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False)
            
            data_hora_atual = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
            nome_arquivo_ceps = f"Planilha_Correios_Finalizada_{data_hora_atual}.xlsx"
            
            st.download_button(
                label="📥 Baixar Planilha Finalizada",
                data=buffer.getvalue(),
                file_name=nome_arquivo_ceps,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
