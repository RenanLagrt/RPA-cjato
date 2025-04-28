import os
import re
import io
import sys
import time
import psutil 
import subprocess
import pytesseract
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
from gerar_cracha import gerar_cracha
from pdf2image import convert_from_path
from datetime import datetime
from dotenv import load_dotenv
from selenium import webdriver
from openpyxl import load_workbook
from itertools import zip_longest
from datetime import datetime, timedelta
from openpyxl.cell.cell import MergedCell
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from Automação_Sertras import AutomaçãoSertras
from Automação_Documentos import AutomaçãoDocumentos

st.set_page_config(layout="wide")

data_atual = datetime.now().strftime("%d-%m-%Y") 

contratos = pd.read_json("info-contratos.json")


if "executado" not in st.session_state:
    st.session_state["executado"] = False  

if "dados_processados" not in st.session_state:
    st.session_state["dados_processados"] = None

# Exibição do cabeçalho
logo_concrejato = r'img\LOGO CONCREJATO.png'

logo_consorcio = r"img\Logo Consórcio.jpg"

col1, col2, col3 = st.columns([1, 4, 1])  

with col1:
    if os.path.exists(logo_concrejato):
        st.image(logo_concrejato, width=220)

with col2:
    st.markdown(
        "<h1 style='text-align: center; color: #004080; font-size: 50px;'>CENTRAL AUTOMAÇÃO QSMS</h1>", 
        unsafe_allow_html=True
    )  

with col3:
    if os.path.exists(logo_consorcio):
        st.image(logo_consorcio, width=220)

# Linha Separadora 
st.markdown("<hr style='border: 1px solid #004080;'>", unsafe_allow_html=True)

placeholder_botao = st.empty()

if not st.session_state["executado"]:
    col_empty1, col_button, col_empty2 = st.columns([2, 2, 2])

    with col_button:
        escolha_contrato = st.selectbox("Selecione um Contrato:", list(contratos.keys()))

    automaçãosertras = AutomaçãoSertras(contratos, escolha_contrato)
    automaçãodocumentos = AutomaçãoDocumentos(contratos, escolha_contrato)

    with col_button:
        botao_relatorio_sertras = st.button("Relatório Sertras", key="baixar_relatório", help="Clique para executar a automação", use_container_width=True, type="primary")
        botao_envio_sertras= st.button("Envio Sertras", key="enviar_documentos", help="Clique para executar a automação", use_container_width=True, type="primary")
        botao_relatorio_documentos = st.button("Relatório Documentos", key="gerar_relatório", help="Clique para executar a automação", use_container_width=True, type="primary")
        botao_gerar_documentos = st.button("Gerar Documentos", key="gerar_documentos", help="Clique para executar a automação", use_container_width=True, type="primary")    
        botao_gerar_cracha = st.button("Gerar Crachá", key="gerar_cracha", help="Clique para executar a automação", use_container_width=True, type="primary")

    col1, col_spinner, col3 = st.columns([100,1,100])

    if botao_relatorio_sertras:
        with col_spinner:
            st.markdown('<div class="spinner-container">', unsafe_allow_html=True)
            with st.spinner(""):
                automaçãosertras.GerarRelatório()

    if botao_envio_sertras: 
        with col_spinner:
            st.markdown('<div class="spinner-container">', unsafe_allow_html=True) 
            with st.spinner(""):
                tabela, documentos_não_encontrados, documentos_encontrados, documentos_enviados, datas_extraidas, vencimentos_projetados, vencimentos_enviados, erro_envio, documentos_atualizados, documentos_nao_atualizados, datas_modificacao = automaçãosertras.EnvioSertras()

                st.session_state["dados_processados"] = {
                    "tabela": tabela,
                    "documentos_não_encontrados": documentos_não_encontrados,
                    "documentos_encontrados": documentos_encontrados,
                    "documentos_nao_atualizados": documentos_nao_atualizados,
                    "documentos_atualizados": documentos_atualizados,
                    "erro_envio": erro_envio,
                    "datas_extraidas": datas_extraidas,
                    "datas_modificacao" : datas_modificacao,
                    "vencimentos_projetados": vencimentos_projetados,
                    "documentos_enviados": documentos_enviados,
                    "vencimentos_enviados": vencimentos_enviados
                }

                st.session_state["executado"] = True
                st.rerun()

    if botao_relatorio_documentos: 
        with col_spinner:
            st.markdown('<div class="spinner-container">', unsafe_allow_html=True) 
            with st.spinner(""):
                automaçãodocumentos.ExibirRelatório()

    if botao_gerar_documentos:
        with col_spinner:
            st.markdown('<div class="spinner-container">', unsafe_allow_html=True) 
            with st.spinner(""):
                try:
                    automaçãodocumentos.GerarDocumentos()

                except FileNotFoundError:
                    automaçãodocumentos.GerarRelatório()
                    automaçãodocumentos.GerarDocumentos()

    if botao_gerar_cracha:
        with col_spinner:
            st.markdown('<div class="spinner-container">', unsafe_allow_html=True) 
            with st.spinner(""):         
                gerar_cracha()

if st.session_state["dados_processados"]:
    dados = st.session_state["dados_processados"]
    
    df_sertras = dados["tabela"]

    df_documentos = pd.DataFrame(list(zip_longest(dados["documentos_não_encontrados"], dados["documentos_encontrados"], dados["datas_modificacao"],
                                        dados["documentos_nao_atualizados"], dados["documentos_atualizados"], fillvalue="---")),

                                columns=["DOCUMENTOS NÃO ENCONTRADOS", "DOCUMENTOS ENCONTRADOS", "DATAS MODIFICAÇÃO", "DOCUMENTOS NÃO ATUALIZADOS", "DOCUMENTOS ATUALIZADOS"])

    df_relatorio = pd.DataFrame(list(zip_longest(dados["erro_envio"], dados["datas_extraidas"], dados["vencimentos_projetados"], 
                                       dados["documentos_enviados"], dados["vencimentos_enviados"], fillvalue="---")),

                                columns=["DOCUMENTOS SEM DATA EXTRAÍDA","DATAS EXTRAÍDAS", "VENCIMENTOS PROJETADOS", "DOCUMENTOS ENVIADOS", "VENCIMENTOS ENVIADOS"])
    
    @st.cache_data
    def to_excel_cached(df, sheet_name):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        return output.getvalue()

    excel_sertras = to_excel_cached(df_sertras, "Pendencias_Sertras")
    excel_documentos = to_excel_cached(df_documentos, "Relacao_Documentos")
    excel_relatorio = to_excel_cached(df_relatorio, "Relatorio_Execução")

    centered_style = [
        {"selector": "thead th", "props": [("background-color", "blue"), ("color", "white"), ("font-weight", "bold"), ("text-align", "center")]},
        {"selector": "tbody td", "props": [("text-align", "center")]}
    ]

    df_sertras_html = df_sertras.style.set_table_styles(centered_style).hide(axis="index").to_html()
    df_documentos_html = df_documentos.style.set_table_styles(centered_style).hide(axis="index").to_html()
    df_relatorio_html = df_relatorio.style.set_table_styles(centered_style).hide(axis="index").to_html()

    def exibir_tabela(titulo, df, arquivo_excel, nome_arquivo):
        col1, col2, col3 = st.columns([0.5, 5, 0.5])  
        with col2:
            st.markdown(f"## 📋 {titulo}")
            st.markdown(df,unsafe_allow_html=True)
        with col3:
            st.download_button(
                data=arquivo_excel,
                label="⬇️",
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        st.markdown("<br><br>", unsafe_allow_html=True)

    exibir_tabela("RELATÓRIO PENDÊNCIAS SERTRAS", df_sertras_html, excel_sertras, f"PENDÊNCIA_SERTRAS {data_atual}.xlsx")
    exibir_tabela("RELAÇÃO DOCUMENTOS", df_documentos_html, excel_documentos, f"RELAÇÃO_DOCUMENTOS {data_atual}.xlsx")
    exibir_tabela("RELATÓRIO EXECUÇÃO", df_relatorio_html, excel_relatorio, f"RELATÓRIO_EXECUÇÃO {data_atual}.xlsx")

streamlit_rodando = False

for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
    try:
        if proc.info['cmdline'] and any("streamlit" in cmd for cmd in proc.info['cmdline']):
            streamlit_rodando = True
            break  
    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
        pass

if not streamlit_rodando:
    subprocess.Popen([sys.executable, "-m", "streamlit", "run", "app.py"], shell=True)
