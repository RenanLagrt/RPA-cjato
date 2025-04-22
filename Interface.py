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

contratos = {
    "OB186 - INHAÚMA" : {
        "botão contrato": "/html/body/div[2]/section/div/div/div[2]/div/div[1]/a/div/div/h5",

        "logo" : os.path.join(os.path.expanduser("~"), "CONSORCIO CONCREJATOEFFICO LOTE 1", "Central de Arquivos - QSMS", "000 ATUAL - OBRA 186 - INHAÚMA", "Logo Consórcio.jpg"),
        
        "documentos" : {
            "DP": ["RG", "CTPS", "CTF", "FRE", "CNH"],
            "QSMS": ["ASO", "EPI", "NR10", "NR11", "NR12", "NR33", "NR35", "DIPLOMA"],
        },
        "documentos/função" : {
            "ELETRICISTA DE REPARO DE REDE DE SANEAMENTO" : [
                "ASO", "FRE", "EPI", "NR6", "NR10", "NR12", "NR18", "NR33", "NR35", "OS"
            ],
            "OPERADOR DE REPARO DE REDE DE SANEAMENTO" : [
                "ASO", "FRE", "EPI", "NR6", "NR12", "NR18", "NR33", "NR35", "OS"
            ],
            "1/2 OFICIAL DE REPARO DE REDE DE SANEAMENTO CIVIL" : [
                "ASO", "FRE", "EPI", "NR6", "NR12", "NR18", "NR33", "NR35", "OS"
            ],
            "AUXILIAR DE REPARO DE REDE DE SANEAMENTO" : [
                "ASO", "FRE", "EPI", "NR6", "NR12", "NR18", "NR33", "NR35", "OS"
            ],
            "ENCARREGADO DE REPARO DE REDE DE SANEAMENTO" : [
                "ASO", "FRE", "EPI", "NR6", "NR18", "NR33", "NR35", "OS"
            ],
            "OPERADOR RETROESCAVADEIRA" : [
                "ASO", "FRE", "EPI", "NR6", "NR11", "NR18", "OS"
            ],
            "ESTAGIARIO" : [
                "ASO", "EPI", "NR6", "NR18", "OS"
            ],
            "OUTRAS": [
                "ASO", "FRE", "EPI", "NR6", "NR18", "OS"
            ]
        },
        "diretorio funcionarios" : {
            "DP": os.path.join(os.path.expanduser("~"), "CONSORCIO CONCREJATOEFFICO LOTE 1", "Central de Arquivos - SERTRAS ARQUIVO PESSOAL"),
            "QSMS": os.path.join(os.path.expanduser("~"), "CONSORCIO CONCREJATOEFFICO LOTE 1", "Central de Arquivos - QSMS", "000 ATUAL - OBRA 186 - INHAÚMA", "Documentação Funcionários"),
        },
        "diretorio efetivo": os.path.join(os.path.expanduser("~"), "CONSORCIO CONCREJATOEFFICO LOTE 1", "Central de Arquivos - QSMS", "000 ATUAL - OBRA 186 - INHAÚMA", "Efetivo", "QUANTITATIVO CONSORCIO.xlsx"),

        "diretorio modelos": os.path.join(os.path.expanduser("~"), "CONSORCIO CONCREJATOEFFICO LOTE 1", "Central de Arquivos - QSMS","000 ATUAL - OBRA 186 - INHAÚMA","MATRIZ DE DOCUMENTOS", "MODELOS"),

        "diretorio saida": os.path.join(os.path.expanduser("~"),"CONSORCIO CONCREJATOEFFICO LOTE 1", "Central de Arquivos - QSMS","000 ATUAL - OBRA 186 - INHAÚMA","MATRIZ DE DOCUMENTOS", "DOCUMENTAÇÃO CRIADA"),

        "mapeamentos":{
            "mapeamento_para_documentos": {
                "ELETRICISTA DE REPARO DE REDE DE SANEAMENTO": {
                    "RG": '//*[@id="edtRequisito_Valor_1"]',
                    "CTPS": '//*[@id="edtRequisito_Valor_2"]',
                    "FRE": '//*[@id="edtRequisito_Valor_3"]',
                    "CTF": '//*[@id="edtRequisito_Valor_4"]',
                    "ASO": '//*[@id="edtRequisito_Valor_6"]',
                    "EPI": '//*[@id="edtRequisito_Valor_8"]',
                    "DIPLOMA": '//*[@id="edtRequisito_Valor_9"]',
                    "NR10": '//*[@id="edtRequisito_Valor_11"]',
                    "NR33": '//*[@id="edtRequisito_Valor_13"]',
                    "NR35": '//*[@id="edtRequisito_Valor_15"]',
                },
                "OPERADOR DE ESCAVADEIRA": {
                    "RG": '//*[@id="edtRequisito_Valor_1"]',
                    "CTPS": '//*[@id="edtRequisito_Valor_2"]',
                    "FRE": '//*[@id="edtRequisito_Valor_3"]',
                    "CTF": '//*[@id="edtRequisito_Valor_4"]',
                    "ASO": '//*[@id="edtRequisito_Valor_6"]',
                    "EPI": '//*[@id="edtRequisito_Valor_8"]',
                    "CNH": '//*[@id="edtRequisito_Valor_10"]',
                    "NR11": '//*[@id="edtRequisito_Valor_12"]',
                },
                "OUTRAS": {
                    "RG": '//*[@id="edtRequisito_Valor_1"]',
                    "CTPS": '//*[@id="edtRequisito_Valor_2"]',
                    "FRE": '//*[@id="edtRequisito_Valor_3"]',
                    "CRF": '//*[@id="edtRequisito_Valor_4"]',
                    "ASO": '//*[@id="edtRequisito_Valor_6"]',
                    "EPI": '//*[@id="edtRequisito_Valor_8"]',
                    "DIPLOMA": '//*[@id="edtRequisito_Valor_9"]',
                    "NR10": '//*[@id="edtRequisito_Valor_11"]',
                    "NR12": '//*[@id="edtRequisito_Valor_10"]',
                    "NR33": '//*[@id="edtRequisito_Valor_12"]',
                    "NR35": '//*[@id="edtRequisito_Valor_14"]',
                },
            },
            "mapeamento_para_datas": {
                "ELETRICISTA DE REPARO DE REDE DE SANEAMENTO": {
                    "RG": '//*[@id="edtRequisito_Valor_0"]',
                    "ASO": '//*[@id="edtRequisito_Valor_5"]',
                    "EPI": '//*[@id="edtRequisito_Valor_7"]',
                    "NR10": '//*[@id="edtRequisito_Valor_10"]',
                    "NR33": '//*[@id="edtRequisito_Valor_12"]',
                    "NR35": '//*[@id="edtRequisito_Valor_14"]',
                },
                "OPERADOR DE ESCAVADEIRA": {
                    "RG": '//*[@id="edtRequisito_Valor_0"]',
                    "ASO": '//*[@id="edtRequisito_Valor_5"]',
                    "EPI": '//*[@id="edtRequisito_Valor_7"]',
                    "CNH": '//*[@id="edtRequisito_Valor_9"]',
                    "NR11": '//*[@id="edtRequisito_Valor_11"]',
                },
                "OUTRAS": {
                    "RG": '//*[@id="edtRequisito_Valor_0"]',
                    "ASO": '//*[@id="edtRequisito_Valor_5"]',
                    "EPI": '//*[@id="edtRequisito_Valor_7"]',
                    "NR11": '//*[@id="edtRequisito_Valor_11"]',
                    "NR12": '//*[@id="edtRequisito_Valor_9"]',
                    "NR33": '//*[@id="edtRequisito_Valor_11"]',
                    "NR35": '//*[@id="edtRequisito_Valor_13"]',
                },
            },
            "mapeamento_para_comentarios": {
                "OUTRAS": {
                    "RG": '//*[@id="edtRequisito_Descricao_1"]',
                    "CTPS": '//*[@id="edtRequisito_Descricao_2"]',
                    "FRE": '//*[@id="edtRequisito_Descricao_3"]',
                    "CTF": '//*[@id="edtRequisito_Descricao_4"]',
                    "ASO": '//*[@id="edtRequisito_Descricao_6"]',
                    "EPI": '//*[@id="edtRequisito_Descricao_8"]'
                }
            }
        }
    },

    "OB201 - SÃO GONÇALO": {
        "botão contrato": "/html/body/div[2]/section/div/div/div[2]/div/div[2]/a/div/div/h5",

        "logo" : os.path.join(os.path.expanduser("~"), "CONSORCIO CONCREJATOEFFICO LOTE 1", "Central de Arquivos - QSMS", "000 ATUAL - OBRA 201 - SÃO GONÇALO", "LOGO CONCREJATO.png"),

        "documentos" : {
            "DP": ["RG", "CTPS", "CTF", "FRE", "CNH"],
            "QSMS": ["ASO", "EPI", "NR10", "NR11", "NR12", "NR33", "NR35", "CERTIFICADO DE CLASSE"],
        },
        "documentos/função" : {
            "ELETRICISTA FORCA E CONTROLE" : [
                "ASO", "FRE", "EPI", "NR6", "NR10", "NR12", "NR18", "NR33", "NR35", "OS"
            ],
            "OPERADOR DE REPARO DE REDE DE SANEAMENTO" : [
                "ASO", "FRE", "EPI", "NR6", "NR12", "NR18", "NR33", "NR35", "OS"
            ],
            "1/2 OFICIAL DE REPARO DE REDE DE SANEAMENTO CIVIL" : [
                "ASO", "FRE", "EPI", "NR6", "NR12", "NR18", "NR33", "NR35", "OS"
            ],
            "AUXILIAR DE REPARO DE REDE DE SANEAMENTO" : [
                "ASO", "FRE", "EPI", "NR6", "NR12", "NR18", "NR33", "NR35", "OS"
            ],
            "ENCARREGADO DE REPARO DE REDE DE SANEAMENTO" : [
                "ASO", "FRE", "EPI", "NR6", "NR18", "NR33", "NR35", "OS"
            ],
            "OPERADOR RETROESCAVADEIRA" : [
                "ASO", "FRE", "EPI", "NR6", "NR11", "NR18", "OS"
            ],
            "ESTAGIARIO" : [
                "ASO", "EPI", "NR6", "NR18", "OS"
            ],
            "OUTRAS": [
                "ASO", "FRE", "EPI", "NR6", "NR18", "OS"
            ]
        },
        "diretorio funcionarios" : {
            "DP": os.path.join(os.path.expanduser("~"), "CONSORCIO CONCREJATOEFFICO LOTE 1", "Central de Arquivos - SERTRAS ARQUIVO PESSOAL"),
            "QSMS": os.path.join(os.path.expanduser("~"),"CONSORCIO CONCREJATOEFFICO LOTE 1","Central de Arquivos - QSMS","000 ATUAL - OBRA 201 - SÃO GONÇALO","DOCUMENTAÇÃO DE FUNCIONÁRIOS","CONCREJATO"),
        },
        "diretorio efetivo": os.path.join(os.path.expanduser("~"), "CONSORCIO CONCREJATOEFFICO LOTE 1", "Central de Arquivos - QSMS", "000 ATUAL - OBRA 201 - SÃO GONÇALO", "EFETIVO", "QUANTITATIVO CONSORCIO.xlsx"),

        "diretorio modelos": os.path.join(os.path.expanduser("~"), "CONSORCIO CONCREJATOEFFICO LOTE 1", "Central de Arquivos - QSMS", "000 ATUAL - OBRA 201 - SÃO GONÇALO", "MATRIZ DE DOCUMENTOS", "MODELOS"),

        "diretorio saida": os.path.join(os.path.expanduser("~"),"CONSORCIO CONCREJATOEFFICO LOTE 1", "Central de Arquivos - QSMS", "000 ATUAL - OBRA 201 - SÃO GONÇALO", "MATRIZ DE DOCUMENTOS", "DOCUMENTAÇÃO CRIADA"),

        "mapeamentos":{
            "mapeamento_para_documentos": {
                "ELETRICISTA FORCA E CONTROLE": {
                    "RG": '//*[@id="edtRequisito_Valor_1"]',
                    "CTPS": '//*[@id="edtRequisito_Valor_2"]',
                    "FRE": '//*[@id="edtRequisito_Valor_3"]',
                    "CTF": '//*[@id="edtRequisito_Valor_4"]',
                    "ASO": '//*[@id="edtRequisito_Valor_6"]',
                    "EPI": '//*[@id="edtRequisito_Valor_8"]',
                    "DIPLOMA": '//*[@id="edtRequisito_Valor_9"]',
                    "CNH": '//*[@id="edtRequisito_Valor_11"]',
                    "NR10": '//*[@id="edtRequisito_Valor_13"]',
                    "NR12": '//*[@id="edtRequisito_Valor_15"]',
                    "NR33": '//*[@id="edtRequisito_Valor_17"]',
                    "NR35": '//*[@id="edtRequisito_Valor_19"]',
                },
                "SOLDADOR": {
                    "RG": '//*[@id="edtRequisito_Valor_1"]',
                    "CTPS": '//*[@id="edtRequisito_Valor_2"]',
                    "FRE": '//*[@id="edtRequisito_Valor_3"]',
                    "CTF": '//*[@id="edtRequisito_Valor_4"]',
                    "ASO": '//*[@id="edtRequisito_Valor_6"]',
                    "EPI": '//*[@id="edtRequisito_Valor_8"]',
                    "DIPLOMA": '//*[@id="edtRequisito_Valor_9"]',
                    "NR12": '//*[@id="edtRequisito_Valor_11"]',
                    "NR33": '//*[@id="edtRequisito_Valor_13"]',
                    "NR35": '//*[@id="edtRequisito_Valor_15"]',
                },
                "ENCARREGADO DE OBRAS": {
                    "RG": '//*[@id="edtRequisito_Valor_1"]',
                    "CTPS": '//*[@id="edtRequisito_Valor_2"]',
                    "FRE": '//*[@id="edtRequisito_Valor_3"]',
                    "CTF": '//*[@id="edtRequisito_Valor_4"]',
                    "ASO": '//*[@id="edtRequisito_Valor_6"]',
                    "EPI": '//*[@id="edtRequisito_Valor_8"]',
                    "DIPLOMA": '//*[@id="edtRequisito_Valor_9"]',
                    "CNH": '//*[@id="edtRequisito_Valor_11"]',
                    "NR33": '//*[@id="edtRequisito_Valor_13"]',
                    "NR35": '//*[@id="edtRequisito_Valor_15"]',
                },
                "COORDENADOR DE OBRAS": {
                    "RG": '//*[@id="edtRequisito_Valor_1"]',
                    "CTPS": '//*[@id="edtRequisito_Valor_2"]',
                    "FRE": '//*[@id="edtRequisito_Valor_3"]',
                    "CTF": '//*[@id="edtRequisito_Valor_4"]',
                    "ASO": '//*[@id="edtRequisito_Valor_6"]',
                    "EPI": '//*[@id="edtRequisito_Valor_8"]',
                    "CNH": '//*[@id="edtRequisito_Valor_10"]',
                    "NR12": '//*[@id="edtRequisito_Valor_12"]',
                },
                "TECNICO DE SEGURANÇA DO TRABALHO": {
                    "RG": '//*[@id="edtRequisito_Valor_1"]',
                    "CTPS": '//*[@id="edtRequisito_Valor_2"]',
                    "FRE": '//*[@id="edtRequisito_Valor_3"]',
                    "CTF": '//*[@id="edtRequisito_Valor_4"]',
                    "ASO": '//*[@id="edtRequisito_Valor_6"]',
                    "EPI": '//*[@id="edtRequisito_Valor_8"]',
                    "DIPLOMA": '//*[@id="edtRequisito_Valor_9"]',
                    "CNH": '//*[@id="edtRequisito_Valor_11"]',
                    "NR33": '//*[@id="edtRequisito_Valor_13"]',
                    "NR35": '//*[@id="edtRequisito_Valor_15"]',
                },
                "SUPERVISOR SEGURANÇA DO TRABALHO": {
                    "RG": '//*[@id="edtRequisito_Valor_1"]',
                    "CTPS": '//*[@id="edtRequisito_Valor_2"]',
                    "FRE": '//*[@id="edtRequisito_Valor_3"]',
                    "CTF": '//*[@id="edtRequisito_Valor_4"]',
                    "ASO": '//*[@id="edtRequisito_Valor_6"]',
                    "EPI": '//*[@id="edtRequisito_Valor_8"]',
                    "DIPLOMA": '//*[@id="edtRequisito_Valor_9"]',
                    "CNH": '//*[@id="edtRequisito_Valor_11"]',
                    "NR33": '//*[@id="edtRequisito_Valor_13"]',
                    "NR35": '//*[@id="edtRequisito_Valor_15"]',
                },
                "OUTRAS": {
                    "RG": '//*[@id="edtRequisito_Valor_1"]',
                    "CTPS": '//*[@id="edtRequisito_Valor_2"]',
                    "FRE": '//*[@id="edtRequisito_Valor_3"]',
                    "CRF": '//*[@id="edtRequisito_Valor_4"]',
                    "ASO": '//*[@id="edtRequisito_Valor_6"]',
                    "EPI": '//*[@id="edtRequisito_Valor_8"]',
                    "DIPLOMA": '//*[@id="edtRequisito_Valor_9"]',
                    "NR10": '//*[@id="edtRequisito_Valor_11"]',
                    "NR12": '//*[@id="edtRequisito_Valor_10"]',
                    "NR33": '//*[@id="edtRequisito_Valor_12"]',
                    "NR35": '//*[@id="edtRequisito_Valor_14"]',
                },
            },
            "mapeamento_para_datas": {
                "ELETRICISTA FORCA E CONTROLE": {
                    "RG": '//*[@id="edtRequisito_Valor_0"]',
                    "ASO": '//*[@id="edtRequisito_Valor_5"]',
                    "EPI": '//*[@id="edtRequisito_Valor_7"]',
                    "CNH": '//*[@id="edtRequisito_Valor_10"]',
                    "NR10": '//*[@id="edtRequisito_Valor_12"]',
                    "NR12": '//*[@id="edtRequisito_Valor_14"]',
                    "NR33": '//*[@id="edtRequisito_Valor_16"]',
                    "NR35": '//*[@id="edtRequisito_Valor_18"]',
                },
                "SOLDADOR": {
                    "RG": '//*[@id="edtRequisito_Valor_0"]',
                    "ASO": '//*[@id="edtRequisito_Valor_5"]',
                    "EPI": '//*[@id="edtRequisito_Valor_7"]',
                    "NR12": '//*[@id="edtRequisito_Valor_10"]',
                    "NR33": '//*[@id="edtRequisito_Valor_12"]',
                    "NR35": '//*[@id="edtRequisito_Valor_14"]',
                },
                "ENCARREGADO DE OBRAS": {
                    "RG": '//*[@id="edtRequisito_Valor_0"]',
                    "ASO": '//*[@id="edtRequisito_Valor_5"]',
                    "EPI": '//*[@id="edtRequisito_Valor_7"]',
                    "CNH": '//*[@id="edtRequisito_Valor_10"]',
                    "NR33": '//*[@id="edtRequisito_Valor_12"]',
                    "NR35": '//*[@id="edtRequisito_Valor_14"]'
                },
                "TECNICO DE SEGURANÇA DO TRABALHO": {
                    "RG": '//*[@id="edtRequisito_Valor_0"]',
                    "ASO": '//*[@id="edtRequisito_Valor_5"]',
                    "EPI": '//*[@id="edtRequisito_Valor_7"]',
                    "CNH": '//*[@id="edtRequisito_Valor_10"]',
                    "NR33": '//*[@id="edtRequisito_Valor_12"]',
                    "NR35": '//*[@id="edtRequisito_Valor_14"]'
                },
                "SUPERVISOR SEGURANÇA DO TRABALHO": {
                    "RG": '//*[@id="edtRequisito_Valor_0"]',
                    "ASO": '//*[@id="edtRequisito_Valor_5"]',
                    "EPI": '//*[@id="edtRequisito_Valor_7"]',
                    "CNH": '//*[@id="edtRequisito_Valor_10"]',
                    "NR33": '//*[@id="edtRequisito_Valor_12"]',
                    "NR35": '//*[@id="edtRequisito_Valor_14"]'
                },
                "OUTRAS": {
                    "RG": '//*[@id="edtRequisito_Valor_0"]',
                    "ASO": '//*[@id="edtRequisito_Valor_5"]',
                    "EPI": '//*[@id="edtRequisito_Valor_7"]',
                    "NR11": '//*[@id="edtRequisito_Valor_11"]',
                    "NR12": '//*[@id="edtRequisito_Valor_9"]',
                    "NR33": '//*[@id="edtRequisito_Valor_11"]',
                    "NR35": '//*[@id="edtRequisito_Valor_13"]',
                },
            },
            "mapeamento_para_comentarios": {
                "OUTRAS": {
                    "RG": '//*[@id="edtRequisito_Descricao_1"]',
                    "CTPS": '//*[@id="edtRequisito_Descricao_2"]',
                    "FRE": '//*[@id="edtRequisito_Descricao_3"]',
                    "CTF": '//*[@id="edtRequisito_Descricao_4"]',
                    "ASO": '//*[@id="edtRequisito_Descricao_6"]',
                    "EPI": '//*[@id="edtRequisito_Descricao_8"]'
                }
            }
        }
    }
}

# ---------------------------|-----------------------------------------|----------------------------|----------------------------|-----------------------------|--------------------

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
        escolha_contrato = st.selectbox("Selecione uma opção:", list(contratos.keys()))

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
                caminho_saida =  f"RELATÓRIO_DOCUMENTAÇÃO {data_atual}.xlsx"
                automaçãodocumentos.ExibirRelatório(caminho_saida)

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
    subprocess.Popen([sys.executable, "-m", "streamlit", "run", "Interface.py"], shell=True)
