import xlwings as xw
import os
import re
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

modelo_cracha = "CRACHA DE SEGURANÇA - CONCREJATO MODELO.xlsx"

base_dir = os.path.expanduser(r"~\CONSORCIO CONCREJATOEFFICO LOTE 1\Central de Arquivos - QSMS\000 ATUAL - OBRA 201 - SÃO GONÇALO")

base_dir_empresas = {
    "MCA CONSTRUÇÕES E REFORMAS": os.path.join(base_dir, r"TERCEIROS\MCA-OB201-SÃO GONÇALO E MARICÁ"),
    "FOCUS ARQUITETURA": os.path.join(base_dir, r"TERCEIROS\FOCUS ARQUITETURA-OB201-SÃO GONÇALO E MARICÁ"),
    "KR LOCAÇÕES": os.path.join(base_dir, r"TERCEIROS\KR LOCAÇÕES-OB201-SÃO GONÇALO E MARICÁ"),
    "UNIK": os.path.join(base_dir, r"TERCEIROS\UNIK MANUTENCAO E LOCACAO-OB201-SÃO GONÇALO E MARICÁ"),
    "UNIMAK": os.path.join(base_dir, r"TERCEIROS\UNIMAK-OB201-SÃO GONÇALO E MARICÁ"),
    "UNIRENT": os.path.join(base_dir, r"TERCEIROS\UNIRENT-OB201-SÃO GONÇALO E MARICÁ"),
}
    
planilha_empresas = {
    "MCA CONSTRUÇÕES E REFORMAS": os.path.join(base_dir_empresas["MCA CONSTRUÇÕES E REFORMAS"], "2025_SEGTRAB_MCA_TECEIROS_ST_CJATO_REV.00.xlsx"),
    "FOCUS ARQUITETURA": os.path.join(base_dir_empresas["FOCUS ARQUITETURA"], "2025_SEGTRAB_FOCUS_TECEIROS_ST_CJATO_REV.00.xlsx"),
    "KR LOCAÇÕES": os.path.join(base_dir_empresas["KR LOCAÇÕES"], "2025_SEGTRAB_KR LOCACOES_ST_CJATO_REV.00.xlsx"),
    "UNIK": os.path.join(base_dir_empresas["UNIK"], "2025_SEGTRAB_UNIK_TECEIROS_ST_CJATO_REV.00.xlsx"),
    "UNIMAK": os.path.join(base_dir_empresas["UNIMAK"], "2025_SEGTRAB_UNIMAK_TECEIROS_ST_CJATO_REV.00.xlsx"),
    "UNIRENT": os.path.join(base_dir_empresas["UNIRENT"], "2025_SEGTRAB_UNIRENT_TECEIROS_ST_CJATO_REV.00.xlsx"),
}

dados = []

for empresa, arquivo in planilha_empresas.items():
    if os.path.exists(arquivo):
        df = pd.read_excel(arquivo, header=14)
        df.columns = df.columns.str.strip().str.replace(r"\s+", " ", regex=True).str.upper()
        df = df.dropna(subset=["NOME DO FUNCIONÁRIO"], how="all")
        # print(f"Colunas encontradas em {empresa}: {list(df.columns)}")

        for _, row in df.iterrows():
            dados.append({
                "EMPRESA": empresa,
                "FUNCIONÁRIO": row.get("NOME DO FUNCIONÁRIO", ""),
                "FUNÇÃO": row.get("FUNÇÃO", ""),
                "CRACHÁ": row.get("CRACHÁ", ""),
                "ASO": row.get("ASO ADMISSIONAL DATA", ""),
                "INTEGRAÇÃO": row.get("DATA DA INTEGRAÇÃO DA AEGEA", ""),
                "NR6": row.get("FICHA DE EPI NR-6 DATA", ""),
                # "NR10": row.get("CERTIFICADO NR - 10 DATA", ""),
                "NR11": row.get("CERTIFICADO NR - 11 DATA", ""),
                "NR12": row.get("CERTIFICADO NR - 12 DATA", ""),
                "NR18": row.get("CERTIFICADO NR - 18 DATA", ""),
                "NR33 SUPERVISOR DE ENTRADA": row.get("VENCIMENTO NR-33 (SUPERVISOR DE ENTRADA)", ""),
                "NR33 TRABALHADOR AUTORIZADO / VIGIA": row.get("VENCIMENTO NR 33 - TRABALHADOR AUTORIZADO / VIGIA", ""),
                "NR35": row.get("CERTIFICADO NR-35 DATA", ""),
                "DIREÇÃO DEFENSIVA": row.get("DIREÇÃO DEFENSIVA DATA", ""),
                "RETROESCAVADEIRA": row.get("CERTIFICADO OPERADOR DE RETRO ESCAVADEIRA DATA", ""),
                "MUNCK": row.get("CERTIFICADO OPERADOR DE GUINDAUTO DATA", ""),
            })

df_final = pd.DataFrame(dados)
df_final.to_excel("dados_gerados.xlsx", index=False)

for col in ["ASO", "INTEGRAÇÃO", "NR6", "NR12", "NR18", "NR35", "DIREÇÃO DEFENSIVA", "RETROESCAVADEIRA", "MUNCK", "NR33 SUPERVISOR DE ENTRADA", "NR33 TRABALHADOR AUTORIZADO / VIGIA"]:
    df_final[col] = pd.to_datetime(df_final[col], errors="coerce")

    df_final[col] = df_final[col].apply(
        lambda x: (pd.to_datetime("1899-12-30") + pd.to_timedelta(x, unit="D")) if isinstance(x, (int, float)) else x)

df_final["ASO"] = df_final["ASO"].apply(lambda x: x + relativedelta(years=1) if pd.notnull(x) else x)

app = xw.App(visible=False)

for _, row in df_final.iterrows():
    if str(row["CRACHÁ"]).strip().upper() != "OK":  
        try:
            wb = app.books.open(modelo_cracha)
            ws = wb.sheets[0]

            foto = os.path.join(base_dir_empresas[row["EMPRESA"]], "2-FUNCIONARIOS", f"{row['FUNCIONÁRIO']}.jpg")

            if os.path.exists(foto):
                ws.shapes["Rectangle 147"].delete()
                ws.pictures.add(foto, name="Rectangle 147", update=True, left=ws.shapes["Rectangle 147"].left, top=ws.shapes["Rectangle 147"].top, width=ws.shapes["Rectangle 147"].width, height=ws.shapes["Rectangle 147"].height)

            ws.shapes["CaixaDeTexto 25"].text = row["FUNCIONÁRIO"]
            ws.shapes["CaixaDeTexto 26"].text = row["FUNÇÃO"]
            ws.shapes["CaixaDeTexto 27"].text = row["ASO"].strftime("%d/%m/%Y") if isinstance(row["ASO"], pd.Timestamp) else ""
            ws.shapes["CaixaDeTexto 28"].text = row["EMPRESA"]

            ws.range("N12").value = row["INTEGRAÇÃO"].date() if isinstance(row["INTEGRAÇÃO"], pd.Timestamp) else ""
            ws.range("L12").value = row["DIREÇÃO DEFENSIVA"].date() if isinstance(row["DIREÇÃO DEFENSIVA"], pd.Timestamp) else ""
            ws.range("J12").value = row["NR35"].date() if isinstance(row["NR35"], pd.Timestamp) else ""
            ws.range("H12").value = row["MUNCK"].date() if isinstance(row["MUNCK"], pd.Timestamp) else ""
            ws.range("G12").value = row["NR33 TRABALHADOR AUTORIZADO / VIGIA"].date() if isinstance(row["NR33 TRABALHADOR AUTORIZADO / VIGIA"], pd.Timestamp) else ""
            ws.range("F12").value = row["NR33 SUPERVISOR DE ENTRADA"].date() if isinstance(row["NR33 SUPERVISOR DE ENTRADA"], pd.Timestamp) else ""
            ws.range("E12").value = row["RETROESCAVADEIRA"].date() if isinstance(row["RETROESCAVADEIRA"], pd.Timestamp) else ""
            ws.range("D12").value = row["NR6"].date() if isinstance(row["NR6"], pd.Timestamp) else ""
            ws.range("C12").value = row["NR12"].date() if isinstance(row["NR12"], pd.Timestamp) else ""
            ws.range("B12").value = row["NR18"].date() if isinstance(row["NR18"], pd.Timestamp) else ""

            nome_arquivo = f"CRACHÁ - {row['FUNCIONÁRIO']}.xlsx"
            wb.save(nome_arquivo)  
            wb.close()

            print(f"Crachá gerado: {nome_arquivo}")

        except Exception as e:
            print(f"Erro ao processar {row['FUNCIONÁRIO']}: {e}")
            continue  

app.quit()



















