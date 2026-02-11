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
    <p class="titulo">💡 SOLUX 2026: Conciliação Automática Real 💡</p>
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
    with st.spinner('O robô SOLUX está instalando as fórmulas mágicas... 🕵️‍♂️✨'):
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
                    banco[f_cod] = {"info": f"{f_cod}", "df": []}
                elif len(lin) >= 10 and pd.notna(lin[0]) and any(x in str(lin[0]) for x in ['/', '-']):
                    deb, cre = to_num(lin[8]), to_num(lin[9])
                    if deb != 0 or cre != 0:
                        hist = str(lin[2]).strip()
                        if 'TOTAL' in hist.upper(): continue
                        
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
                    f_cab = wb.add_format({'bold': 1, 'bg_color': '#9B8ADE', 'font_color': 'white', 'border': 1, 'align': 'center'})
                    f_m = wb.add_format({'num_format': '#,##0.00', 'border': 1})
                    f_c = wb.add_format({'border': 1, 'align': 'center'})

                    for cod, conteudo in banco.items():
                        df = pd.DataFrame(conteudo["df"])
                        if df.empty: continue
                        
                        ws = wb.add_worksheet(str(cod)[:31])
                        ws.set_column('B:G', 18); ws.set_column('C:C', 40); ws.set_column('I:L', 18)
                        
                        # Escreve o Razão
                        headers = ["Data", "Histórico", "NF_Ajustada", "Débito", "Crédito"]
                        for ci, v in enumerate(headers): ws.write(5, ci+1, v, f_cab)
                        
                        for ri, r in enumerate(df.values):
                            ws.write(6+ri, 1, r[0], f_c)
                            ws.write(6+ri, 2, r[1], f_c)
                            ws.write(6+ri, 3, r[2], f_c) # Esta é a coluna que você vai editar no Excel
                            ws.write_number(6+ri, 4, r[3], f_m)
                            ws.write_number(6+ri, 5, r[4], f_m)

                        # A MÁGICA: TABELA DE CONCILIAÇÃO COM FÓRMULAS
                        nfs_unicas = df["NF_Ajustada"].unique()
                        ws.write(5, 8, "NF (Resumo)", f_cab)
                        ws.write(5, 9, "Soma Débito", f_cab)
                        ws.write(5, 10, "Soma Crédito", f_cab)
                        ws.write(5, 11, "Diferença", f_cab)

                        for ri, nf in enumerate(nfs_unicas):
                            row_idx = 7 + ri
                            ws.write(row_idx-1, 8, nf, f_c)
                            
                            # Fórmulas SUMIF (SOMASE) - Elas olham para a coluna D (NF_Ajustada) e somam E (Deb) e F (Cre)
                            # Se você mudar a NF na coluna D, o resultado aqui muda na hora!
                            razao_nf_range = f"$D$7:$D${6+len(df)}"
                            ws.write_formula(row_idx-1, 9, f'=SUMIF({razao_nf_range}, I{row_idx}, $E$7:$E${6+len(df)})', f_m)
                            ws.write_formula(row_idx-1, 10, f'=SUMIF({razao_nf_range}, I{row_idx}, $F$7:$F${6+len(df)})', f_m)
                            ws.write_formula(row_idx-1, 11, f'=J{row_idx}+K{row_idx}', f_m)

                st.success("✅ Excel 'Vivo' gerado com sucesso!")
                st.download_button("📥 Baixar Relatório", out.getvalue(), "solux_viva.xlsx")
        except Exception as e:
            st.error(f"Erro: {e}")
