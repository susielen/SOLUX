import streamlit as st
import pandas as pd
import re
from io import BytesIO

# 1. Configuração da Página
st.set_page_config(page_title="SOLUX 2026", page_icon="💡", layout="wide")

# 2. ESTILO SOLUX FINAL
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;800&display=swap');
    .stApp { background-color: #F3F0FF; background-image: url("https://www.transparenttextures.com/patterns/cubes.png"); background-attachment: fixed; }
    header[data-testid="stHeader"], [data-testid="stSidebar"] { background-color: #9B8ADE !important; }
    .titulo { font-family: 'Montserrat', sans-serif; color: #4B0082; font-size: 28px; font-weight: 800; text-align: center; padding: 10px; background-color: rgba(230, 224, 255, 0.9); border-radius: 10px; border: 1px solid #9B8ADE; margin-top: -35px; margin-bottom: 25px; }
    [data-testid="stFileUploaderDropzone"] { background-color: rgba(255, 255, 255, 0.6) !important; border: 2px dashed #9B8ADE !important; }
    [data-testid="stFileUploaderDropzone"] button { background-color: #E6E0FF !important; color: #4B0082 !important; border: 1px solid #9B8ADE !important; font-weight: bold !important; }
    [data-testid="stSidebar"] * { color: #FFFFFF !important; font-weight: 600 !important; }
    .stDownloadButton button { background-color: #9B8ADE !important; color: white !important; border-radius: 8px !important; }
    </style>
    <p class="titulo">💡 SOLUX 2026: Conciliação Dinâmica 💡</p>
    """, unsafe_allow_html=True)

def to_num(val):
    try:
        if pd.isna(val) or str(val).strip() == '': return 0.0
        s = str(val).replace('.', '').replace(',', '.')
        return float(re.sub(r'[^-0-9.]', '', s))
    except: return 0.0

with st.sidebar:
    st.header("⚙️ Painel de Controle")
    tipo_robo = st.radio("Este projeto é de:", ["Cliente", "Fornecedor"])
    arquivo = st.file_uploader("Suba o arquivo aqui", type=["xlsx", "xls", "csv"])

if arquivo:
    with st.spinner('O robô SOLUX está organizando as colunas... 🕵️‍♂️✨'):
        try:
            if arquivo.name.endswith('.csv'):
                df_bruto = pd.read_csv(arquivo, header=None, sep=None, engine='python', encoding='latin-1')
            else:
                df_bruto = pd.read_excel(arquivo, header=None)

            nome_emp = "EMPRESA"
            for i in range(min(15, len(df_bruto))):
                if "Empresa:" in str(df_bruto.iloc[i, 0]):
                    nome_emp = str(df_bruto.iloc[i, 2]); break

            banco = {}
            for i in range(len(df_bruto)):
                lin = df_bruto.iloc[i]
                if "Conta:" in str(lin[0]):
                    f_cod = str(lin[1]).strip()
                    dados = []
                    banco[f_cod] = {"info": f"{f_cod} - {str(lin[5]) if len(lin) > 5 else str(lin[2])}", "df": []}
                elif len(lin) >= 10 and pd.notna(lin[0]) and any(x in str(lin[0]) for x in ['/', '-']):
                    deb, cre = to_num(lin[8]), to_num(lin[9])
                    if deb != 0 or cre != 0:
                        hist = str(lin[2]).strip()
                        if 'TOTAL' in hist.upper(): continue
                        
                        # BUSCA DE NF
                        h_up = hist.upper()
                        pats = [r'SERVIÇO\s?PRESTADO\s?(\d+)', r'NF\s?DE\s?S\s?(\d+)', r'FRETE\s?TOMADO\s?(\d+)', r'CTE\s?(\d+)', r'NFE\s?(\d+)', r'SAÍDA\s?(\d+)', r'NF\s?(\d+)']
                        nf_res = "S/N"
                        for p in pats:
                            m = re.findall(p, h_up)
                            if m: nf_res = m[0]; break
                        
                        v_deb, v_cre = (-deb, cre) if tipo_robo == "Fornecedor" else (deb, -cre)
                        banco[f_cod]["df"].append({"Data": str(lin[0]), "Histórico": hist, "NF_Ajustada": nf_res, "Débito": v_deb, "Crédito": v_cre})

            if banco:
                out = BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    wb = writer.book
                    f_cab = wb.add_format({'bold': 1, 'bg_color': '#F2F2F2', 'align': 'center', 'border': 1})
                    f_m = wb.add_format({'num_format': '#,##0.00', 'border': 1})
                    f_c = wb.add_format({'align': 'center', 'border': 1})

                    for cod, conteudo in banco.items():
                        df = pd.DataFrame(conteudo["df"])
                        if df.empty: continue
                        
                        aba = str(cod)[:31]
                        ws = wb.add_worksheet(aba)
                        ws.set_column('B:G', 18); ws.set_column('C:C', 40)
                        
                        # 1. Escreve o RAZÃO com a nova coluna NF_Ajustada
                        ws.write('B2', f"EMPRESA: {nome_emp} ({tipo_robo})", wb.add_format({'bold':1}))
                        headers = ["Data", "Histórico", "NF_Ajustada", "Débito", "Crédito"]
                        for ci, v in enumerate(headers): ws.write(5, ci+1, v, f_cab)
                        
                        for ri, r in enumerate(df.values):
                            ws.write(6+ri, 1, r[0], f_c)
                            ws.write(6+ri, 2, r[1], f_c)
                            ws.write(6+ri, 3, r[2], f_c) # NF_Ajustada
                            ws.write_number(6+ri, 4, r[3], f_m)
                            ws.write_number(6+ri, 5, r[4], f_m)

                        # 2. Escreve a CONCILIAÇÃO (Tabela Dinâmica Manual) ao lado
                        res = df.groupby("NF_Ajustada").agg({"Débito":"sum", "Crédito":"sum"}).reset_index()
                        res["Diferença"] = res["Débito"] + res["Crédito"]
                        
                        ws.write(5, 8, "NF (Conciliação)", f_cab)
                        ws.write(5, 9, "Soma Débito", f_cab)
                        ws.write(5, 10, "Soma Crédito", f_cab)
                        ws.write(5, 11, "Diferença", f_cab)

                        for ri, r in enumerate(res.values):
                            ws.write(6+ri, 8, str(r[0]), f_c)
                            ws.write_number(6+ri, 9, r[1], f_m)
                            ws.write_number(6+ri, 10, r[2], f_m)
                            ws.write_number(6+ri, 11, r[3], f_m)

                st.success("✅ Razão e Tabela de Conciliação gerados!")
                st.download_button("📥 Baixar Relatório", out.getvalue(), "solux_dinamico.xlsx")
        except Exception as e:
            st.error(f"Erro: {e}")
