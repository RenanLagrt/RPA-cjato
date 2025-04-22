import os
import re 
import subprocess
import locale
import shutil
import zipfile
import pandas as pd
from docx import Document
from datetime import datetime
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


locale.setlocale(locale.LC_TIME, "pt_BR.utf8")

class AutomaçãoDocumentos():

    def __init__(self, contratos, contrato_selecionado):
        self.contratos = contratos
        self.contrato_selecionado = contrato_selecionado
    
    def get_info_contrato(self, chamado=None):
        logo = self.contratos[self.contrato_selecionado]["logo"]
        diretorio_efetivo = self.contratos[self.contrato_selecionado]["diretorio efetivo"]
        diretorio_funcionarios = self.contratos[self.contrato_selecionado]["diretorio funcionarios"]["QSMS"]
        diretorio_modelos = self.contratos[self.contrato_selecionado]["diretorio modelos"]
        diretorio_saidas = self.contratos[self.contrato_selecionado]["diretorio saida"]
        documentos_por_função = self.contratos[self.contrato_selecionado]["documentos/função"]

        if chamado == 'RelatórioDocumentos':
            return logo, diretorio_efetivo, diretorio_funcionarios, documentos_por_função
        
        if chamado == "GerarDocumentos":
            return diretorio_modelos, diretorio_saidas
    
    @staticmethod
    def get_diretorio_funcionario(diretorio_funcionarios, nome):
        return os.path.join(os.path.expanduser("~"), diretorio_funcionarios, nome)

    @staticmethod
    def formatar_cpf(cpf):
        cpf = ''.join(filter(str.isdigit, str(cpf)))
        cpf = cpf.zfill(11)
        return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
    
    @staticmethod
    def formatar_data(data):
        return datetime.strftime(data, '%d/%m/%Y')
    
    @staticmethod
    def obter_documentos_requeridos(funcao, documentos_por_funcao):
        return documentos_por_funcao.get(funcao, documentos_por_funcao["OUTRAS"])

    def personalizar_planilha(self, ws, logo):
        fundo_azul = PatternFill(start_color="003399", end_color="003399", fill_type="solid")
        fonte_branca = Font(size=12, color="FFFFFF", bold=True)
        verde_claro = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  
        vermelho = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

        alinhamento_esquerda = Alignment(horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True)
        alinhamento_central = Alignment(horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=True)

        ws.insert_rows(1)

        img = Image(logo)
        img.width = 250
        img.height = 70
        ws.add_image(img, "A1")

        ws["A1"] = "CONTROLE DE DOCUMENTAÇÃO FUNCIONÁRIOS"
        ws["A1"].font = Font(bold=True, size=22, color="003399")
        ws["A1"].alignment = alinhamento_central

        ws.row_dimensions[1].height = 70
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ws.max_column) 

        borda = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )
        ws.row_dimensions[2].height = 34

        for cell in ws[2]:
            cell.fill = fundo_azul
            cell.font = fonte_branca
            cell.alignment = alinhamento_central
            cell.border = borda

        for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
            ws.row_dimensions[row[0].row].height = 23
            for idx, cell in enumerate(row):
                cell.border = borda
                if idx == 0:  
                    cell.alignment = alinhamento_esquerda

                else:  
                    cell.alignment = alinhamento_central

                if cell.value == "OK":
                    cell.fill = verde_claro

                elif cell.value == "P":
                    cell.fill = vermelho

        self.ajustar_largura_colunas(ws)

        ws.freeze_panes = "D3"
        ws.auto_filter.ref = "A2:O{}".format(ws.max_row)

    @staticmethod
    def ajustar_largura_colunas(ws):
        tamanho_minimo_documentos_pendentes = 28
        tamanho_minimo_documentos = 8
        tamanho_minimo_nome = 25

        for idx, column_cells in enumerate(ws.columns):
            max_length = max(len(str(cell.value)) for cell in column_cells if cell.value)
            valid_cells = [cell for cell in column_cells if not isinstance(cell, MergedCell) and cell.value]
            if not valid_cells:
                continue
            column = valid_cells[0].column_letter

            if idx == 0:  
                ws.column_dimensions[column].width = max(max_length + 2, tamanho_minimo_nome)

            elif idx == 15: 
                ws.column_dimensions[column].width = max(max_length + 2, tamanho_minimo_documentos_pendentes)

            elif idx >= 4:  
                ws.column_dimensions[column].width = max(max_length + 2, tamanho_minimo_documentos)

            else:
                ws.column_dimensions[column].width = max_length + 2

    @staticmethod
    def tratar_tabela(tabela_dados):
        tabela_dados["DESC FUNÇÃO"] = tabela_dados["DESC FUNÇÃO"].replace({"1/2 OFICIAL DE REPARO DE REDE DE SANEAMENTO CIVIL" : "MEIO OFICIAL DE REPARO DE REDE DE SANEAMENTO CIVIL",
                        "ASSISTENTE PLANEJAMENTO II" : "ASSISTENTE PLANEJAMENTO",
                        "ASSISTENTE PLANEJAMENTO III" : "ASSISTENTE PLANEJAMENTO",
                        "AUXILIAR ADMINISTRATIVO I" : "AUXILIAR ADMINISTRATIVO",
                        "AUXILIAR ADMINISTRATIVO III" : "AUXILIAR ADMINISTRATIVO",
                        "ELETRICISTA DE REPARO DE REDE DE SANEAMENTO I" : "ELETRICISTA DE REPARO DE REDE DE SANEAMENTO",
                        "ENCARREGADO DE REPARO DE REDE DE SANEAMENTO I" : "ENCARREGADO DE REPARO DE REDE DE SANEAMENTO",
                        "ENCARREGADO DE REPARO DE REDE DE SANEAMENTO II" : "ENCARREGADO DE REPARO DE REDE DE SANEAMENTO"
                        })
        
        return tabela_dados

    def gerar_dados_planilha(self,diretorio_funcionarios, tabela_dados, documentos_por_função, ordem_documentos):
        dados_planilha = []
        
        for nome, linha in tabela_dados.groupby("NOME"):
            caminho_funcionario = self.get_diretorio_funcionario(diretorio_funcionarios, nome)
            funcao = str(linha["DESC FUNÇÃO"].iloc[0])
            admissao = self.formatar_data(linha["DATA ADMISSAO"].iloc[0])
            cpf = self.formatar_cpf(linha["CPF"].iloc[0])

            documentos_requeridos = self.obter_documentos_requeridos(funcao, documentos_por_função)
            documentos_na_pasta = os.listdir(caminho_funcionario) if os.path.isdir(caminho_funcionario) else []
            
            linha_dados = [nome, funcao, cpf, admissao]
            documentos_pendentes = []

            for documento in ordem_documentos:
                if documento in documentos_requeridos:
                    nome_esperado = f"{documento} - {nome}.pdf"
                    linha_dados.append("OK" if nome_esperado in documentos_na_pasta else "P")

                    if nome_esperado not in documentos_na_pasta:
                        documentos_pendentes.append(documento)
                else:
                    linha_dados.append("N/A")
            
            linha_dados.append(" - ".join(documentos_pendentes) if documentos_pendentes else "---")
            dados_planilha.append(linha_dados)

        return dados_planilha

    def GerarRelatório(self):
        wb = Workbook()
        ordem_documentos = ["ASO", "FRE", "EPI", "NR6", "NR10", "NR11", "NR12", "NR18", "NR33", "NR35", "OS"]
        
        for contrato, _ in self.contratos.items():
            logo, diretorio_efetivo, diretorio_funcionarios, documentos_por_função = self.get_info_contrato(contrato,'RelatórioDocumentos')

            tabela_dados = pd.read_excel(diretorio_efetivo)
            tabela_dados = self.tratar_tabela(tabela_dados)

            dados_planilha = self.gerar_dados_planilha(diretorio_funcionarios, tabela_dados, documentos_por_função, ordem_documentos)
            colunas = ["FUNCIONÁRIO", "FUNÇÃO", "CPF", "ADMISSÃO"] + ordem_documentos + ["DOCUMENTAÇÃO PENDENTE"]
            df = pd.DataFrame(dados_planilha, columns=colunas)

            ws = wb.create_sheet(title=contrato)  
            for r_idx, row in enumerate([colunas] + df.values.tolist(), 1):
                for c_idx, value in enumerate(row, 1):
                    ws.cell(row=r_idx, column=c_idx, value=value)

            self.personalizar_planilha(ws,logo)
        
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])
        
        data_atual = datetime.now().strftime("%d-%m-%Y")
        caminho_saida =  f"RELATÓRIO_DOCUMENTAÇÃO {data_atual}.xlsx"
        wb.save(caminho_saida)

    def ExibirRelatório(self, caminho_saida):
        self.GerarRelatório()
        subprocess.run(["cmd", "/c", "start", "", caminho_saida], shell=True)


# --------------------------|----------------------------------|-----------------------------------------------|---------------------------------
        
    def get_modelo(self, documento, funcao, contrato, diretorio_modelos):      
        if not diretorio_modelos:
            print(f"[ERRO] Diretório de modelos não encontrado para o contrato: {contrato}")
            return None
        
        mapa_modelos = {
            doc: os.path.join(diretorio_modelos, "NRs - MODELOS", f"{doc} - MODELO.docx")
            for doc in ["NR6", "NR12", "NR18", "NR33", "NR35"]
        }
        mapa_modelos["OS"] = os.path.join(diretorio_modelos, "OS - MODELOS", f"OS - {funcao}.docx")
        
        modelo = mapa_modelos.get(documento)
        
        if not modelo or not os.path.exists(modelo):
            print(f"[ERRO] Modelo não encontrado ou inexistente para {documento} ({funcao}) no contrato {contrato}.")
            return None
        
        return modelo

    @staticmethod
    def substituir_texto_docx(nome_modelo, substituicoes, diretorio_saida):
        temp_zip_path = diretorio_saida.replace(".docx", "_temp.zip")
        temp_folder = diretorio_saida.replace(".docx", "_temp")
        
        shutil.copy2(nome_modelo, diretorio_saida)
        with zipfile.ZipFile(diretorio_saida, 'r') as docx_zip:
            docx_zip.extractall(temp_folder)
        
        xml_path = os.path.join(temp_folder, "word", "document.xml")
        with open(xml_path, "r", encoding="utf-8") as file:
            xml_content = file.read()
        
        for marcador, novo_texto in substituicoes.items():
            xml_content = re.sub(re.escape(marcador) + r"(\s*</w:t>\s*<w:t[^>]*>)?", novo_texto, xml_content)
        
        with open(xml_path, "w", encoding="utf-8") as file:
            file.write(xml_content)
        
        with zipfile.ZipFile(temp_zip_path, 'w', zipfile.ZIP_DEFLATED) as docx_zip:
            for root, _, files in os.walk(temp_folder):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_folder)
                    docx_zip.write(file_path, arcname)
        
        os.replace(temp_zip_path, diretorio_saida)
        shutil.rmtree(temp_folder, ignore_errors=True)

    def gerar_documentos_pendentes(self,contrato, nome_funcionario, funcao, cpf, admissao, documentos_pendentes, diretorio_modelos, diretorio_saidas):
        for documento in documentos_pendentes:
            modelo = self.get_modelo(documento, funcao, contrato, diretorio_modelos)

            if not modelo:
                print(f"[AVISO] Documento {documento} não criado para {nome_funcionario} ({funcao}) no contrato {contrato}")
                continue  
            
            caminho_saida = f"{diretorio_saidas}/{documento}/{documento} - {nome_funcionario}.docx"
            admissao_formatada = datetime.strptime(admissao, "%d/%m/%Y").strftime("%d de %B de %Y") if documento.startswith("NR") else admissao
            
            substituicoes = {
                "{{NOME}}": nome_funcionario,
                "{{FUNÇÃO}}": funcao,
                "{{CPF}}": cpf,
                "{{ADMISSÃO}}": admissao_formatada,
                "{{TREINAMENTO}}": admissao_formatada
            }
            
            self.substituir_texto_docx(modelo, substituicoes, caminho_saida)
    
    def GerarDocumentos(self):
        data_atual = datetime.now().strftime("%d-%m-%Y")
        diretorio_tabela = f"RELATÓRIO_DOCUMENTAÇÃO {data_atual}.xlsx"
        tabelas_documentacao = pd.read_excel(diretorio_tabela, sheet_name=None)
        
        for contrato, tabela_documentacao in tabelas_documentacao.items():
            diretorio_modelos,  diretorio_saidas = self.get_info_contrato(contrato,"GerarDocumentos")
            tabela_documentacao = pd.read_excel(diretorio_tabela, sheet_name=contrato, header=1)
            
            for _, row in tabela_documentacao.iterrows():
                nome_funcionario = row["FUNCIONÁRIO"]
                funcao = row["FUNÇÃO"]
                cpf = row["CPF"]
                admissao = row["ADMISSÃO"]
                documentos_pendentes = [doc for doc in tabela_documentacao.columns[4:] if row[doc] == "P"]
                
                if documentos_pendentes:
                    self.gerar_documentos_pendentes(contrato, nome_funcionario, funcao, cpf, admissao, documentos_pendentes,diretorio_modelos, diretorio_saidas)

