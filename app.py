import streamlit as st
import pandas as pd
import re
from io import BytesIO

# 1. Configuração da Página
st.set_page_config(page_title="SOLUX", page_icon="💡", layout="wide")

# 2. ESTILO SOLUX (O Streamlit continua com o tema da marca, mas o Excel volta ao original)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;800&display=swap');
    .stApp { background-color: #F3F0FF; }
    header[data-testid="stHeader"], [data-testid="stSidebar"] { background-color: #9B8ADE !important; }
    .titulo { font-family: 'Montserrat', sans-serif; color: #4B0082; font-size: 28px; font-weight: 800; text-align: center; padding: 10px; background-color: rgba(230, 224, 255, 0.9); border-radius: 10px; border: 1px solid #9B8ADE; margin-top: -35px; margin-bottom: 25px; }
    .stDownloadButton button { background-color: #9B8ADE !important; color: white !important; font-weight: bold; width: 100%; height: 50px; }
    </style>
    <p class="titulo">💡 SOLUX: Seu parceiro na conciliação 💡</p>
    """, unsafe_allow_html=True)

def to_num(val):
    try:
        if pd.isna(val) or str(val).strip() == '': return 0.0
        s = str(val).replace('.', '').replace(',', '.')
        return float(re.sub(r'[^-0-9.]', '', s))
    except: return 0.0

with st.sidebar:
    st.header("⚙️ Painel de Controle")
    tipo_robo = st.radio("Este projeto é de:", ["Clientes", "Fornecedores"])
    arquivo = st.file_uploader("Suba o arquivo aqui", type=["xlsx", "xls", "csv"])

if arquivo:
    with st.spinner('Restaurando suas cores favoritas... 🕵️‍♂️✨'):
        try:
            if arquivo.name.endswith('.csv'):
                df_bruto = pd.read_csv(arquivo, header=None, sep=None, engine='python', encoding='latin-1')
            else:
                df_bruto = pd.read_excel(arquivo, header=None)

            nome_emp = "EMPRESA"
            for i in range(min(15, len(df_bruto))):
                if "Empresa:" in str(df_bruto.iloc[i, 0]):
                    nome_emp = str(df_bruto.iloc[i, 2]); break

            banco, f_info = {}, {}
            f_cod, dados = None, []

            for i in range(len(df_bruto)):
                lin = df_bruto.iloc[i]
                if "Conta:" in str(lin[0]):
                    if f_cod and dados: banco[f_cod] = pd.DataFrame(dados)
                    f_cod = str(lin[1]).strip()
                    f_info[f_cod] = f"{f_cod} - {str(lin[5]) if len(lin) > 5 and pd.notna(lin[5]) else str(lin[2])}"
                    dados = []
                elif len(lin) >= 10 and pd.notna(lin[0]) and any(x in str(lin[0]) for x in ['/', '-']):
                    deb, cre = to_num(lin[8]), to_num(lin[9])
                    if deb != 0 or cre != 0:
                        hist = str(lin[2]).strip()
                        if 'TOTAL' in hist.upper(): continue
                        
                        try: data_formatada = pd.to_datetime(lin[0]).strftime('%d/%m/%Y')
                        except: data_formatada = str(lin[0])

                        h_up = hist.upper()
                        # Inclui termos SAÍDA e PRESTADO [2026-02-05]
                        pats = [r'SAÍDA\s?(\d+)', r'PRESTADO\s?(\d+)', r'NF\s?DE\s?S\s?(\d+)', r'NFE\s?(\d+)', r'NF\s?(\d+)']
                        nf_res = None
                        for p in pats:
                            m = re.findall(p, h_up)
                            if m: nf_res = m[0]; break
                        
                        nf = nf_res if nf_res else "S/ N° NF"
                        
                        # Regra de sinais [2026-01-30]
                        if tipo_robo == "Fornecedores": v_deb, v_cre = -deb, cre
                        else: v_deb, v_cre = deb, -cre
                            
                        dados.append({"Data": data_formatada, "NF": nf, "Hist": hist, "Deb": v_deb, "Cred": v_cre, "Aviso": (nf == "S/ N° NF")})

            if f_cod and dados: banco[f_cod] = pd.DataFrame(dados)

            if banco:
                out = BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    wb = writer.book
                    # --- CORES ORIGINAIS RESTAURADAS ---
                    f_cab = wb.add_format({'bold': 1, 'bg_color': '#F2F2F2', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                    f_emp = wb.add_format({'bold': 1, 'font_size': 14, 'align': 'center', 'bg_color': '#D3D3D3', 'border': 1})
                    f_c = wb.add_format({'align': 'center', 'border': 1})
                    f_m = wb.add_format({'num_format': '#,##0.00', 'border': 1})
                    f_s = wb.add_format({'border': 1})
                    f_ama_c = wb.add_format({'align': 'center', 'border': 1, 'bg_color': '#FFFF99'}) # Amarelo de aviso
                    f_vde = wb.add_format({'num_format': '#,##0.00', 'font_color': 'green', 'bold': 1, 'border': 1})
                    f_vrm = wb.add_format({'num_format': '#,##0.00', 'font_color': 'red', 'bold': 1, 'border': 1})

                    for cod, df_emp in banco.items():
                        ws = wb.add_worksheet(str(cod)[:31])
                        ws.hide_gridlines(2)
                        
                        # Tira os triângulos verdes
                        ws.ignore_errors({'number_stored_as_text': 'B1:M1000'})
                        
                        # Coluna A bem fininha (largura que era)
                        ws.set_column('A:A', 2.14) 
                        ws.set_column('B:C', 15); ws.set_column('D:D', 45); ws.set_column('E:F', 18)
                        ws.set_column('G:H', 2.14) 
                        ws.set_column('I:M', 18)
                        
                        ws.merge_range('B2:M2', f"EMPRESA: {nome_emp} ({tipo_robo})", f_emp)
                        ws.merge_range('B4:F4', f_info[cod], f_cab)
                        ws.merge_range('I4:M4', "CONCILIAÇÃO POR NOTA", f_cab)

                        # Razão
                        for ci, v in enumerate(["Data","NF","Histórico","Débito","Crédito"]): ws.write(5, ci+1, v, f_cab)
                        row_f = 5
                        for ri, r in enumerate(df_emp.values):
                            fmt = f_ama_c if r[5] else f_c
                            ws.write(6+ri, 1, r[0], fmt); ws.write(6+ri, 2, r[1], fmt)
                            ws.write(6+ri, 3, r[2], f_s
