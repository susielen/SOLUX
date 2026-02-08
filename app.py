import streamlit as st
import pandas as pd
import re
from io import BytesIO

# 1. Configuração da Página
st.set_page_config(page_title="SOLUX", page_icon="💡", layout="wide")

# 2. ESTILO SOLUX (Layout Lilás e Moderno)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;800&display=swap');
    .stApp { background-color: #F3F0FF; background-image: url("https://www.transparenttextures.com/patterns/cubes.png"); background-attachment: fixed; }
    header[data-testid="stHeader"], [data-testid="stSidebar"] { background-color: #9B8ADE !important; }
    .titulo { font-family: 'Montserrat', sans-serif; color: #4B0082; font-size: 28px; font-weight: 800; text-align: center; padding: 10px; background-color: rgba(230, 224, 255, 0.9); border-radius: 10px; border: 1px solid #9B8ADE; margin-top: -35px; margin-bottom: 25px; }
    
    /* Ajuste na cor e estilo do Painel de Upload */
    [data-testid="stFileUploaderDropzone"] {
        background-color: rgba(255, 255, 255, 0.6) !important;
        border: 2px dashed #9B8ADE !important;
        border-radius: 12px !important;
    }
    
    [data-testid="stSidebar"] * { color: #FFFFFF !important; font-weight: 600 !important; }
    .stDownloadButton button { background-color: #9B8ADE !important; color: white !important; border-radius: 8px !important; }
    </style>
    <p class="titulo">💡 SOLUX: Conciliação Inteligente 💡</p>
    """, unsafe_allow_html=True)

def to_num(val):
    try:
        if pd.isna(val) or str(val).strip() == '': return 0.0
        return float(str(val).replace('.', '').replace(',', '.'))
    except: return 0.0

with st.sidebar:
    st.header("⚙️ Painel de Controle")
    tipo_robo = st.radio("Este projeto é de:", ["Cliente", "Fornecedor"])
    arquivo = st.file_uploader("Suba o arquivo aqui", type=["xlsx", "xls", "csv"])

if arquivo:
    with st.spinner('O robô está finalizando os detalhes... 🕵️‍♂️✨'):
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
                    f_info[f_cod] = f"{f_cod} - {str(lin[5]) if pd.notna(lin[5]) else str(lin[2])}"
                    dados = []
                elif len(lin) > 9:
                    deb, cre = to_num(lin[8]), to_num(lin[9])
                    hist = str(lin[2]).strip()
                    if (deb != 0 or cre != 0) and pd.notna(lin[0]):
                        if 'TOTAL' in hist.upper(): continue
                        try: dt = pd.to_datetime(lin[0]).strftime('%d/%m/%Y')
                        except: dt = str(lin[0])

                        h_up = hist.upper()
                        pats = [r'SERVIÇO\s?PRESTADO\s?(\d+)', r'NF\s?DE\s?S\s?(\d+)', r'FRETE\s?TOMADO\s?(\d+)', r'CTE\s?(\d+)', r'NFE\s?(\d+)', r'SAÍDA\s?(\d+)', r'NF\s?(\d+)']
                        nf_res = None
                        for p in pats:
                            m = re.findall(p, h_up)
                            if m: nf_res = m[0]; break
                        
                        nf = nf_res if nf_res else "S/ N° NF"
                        
                        if tipo_robo == "Fornecedor": v_deb, v_cre = -deb, cre
                        else: v_deb, v_cre = deb, -cre
                        
                        dados.append({"Data": dt, "NF": nf, "Hist": hist, "Deb": v_deb, "Cred": v_cre, "Aviso": (nf == "S/ N° NF")})

            if f_cod and dados: banco[f_cod] = pd.DataFrame(dados)

            if banco:
                out = BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    wb = writer.book
                    f_cab = wb.add_format({'bold': 1, 'bg_color': '#F2F2F2', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                    f_emp = wb.add_format({'bold': 1, 'font_size': 14, 'align': 'center', 'bg_color': '#D3D3D3', 'border': 1})
                    
                    # Formatos normais (Ajustados para evitar o triângulo verde)
                    f_c = wb.add_format({'align': 'center', 'border': 1})
                    f_m = wb.add_format({'num_format': '#,##0.00', 'border': 1})
                    f_s = wb.add_format({'border': 1})
                    
                    # Formatos Amarelo (#FFFF99)
                    cor_ama = '#FFFF99'
                    f_ama_c = wb.add_format({'align': 'center', 'border': 1, 'bg_color': cor_ama})
                    f_ama_m = wb.add_format({'num_format': '#,##0.00', 'border': 1, 'bg_color': cor_ama})
                    f_ama_s = wb.add_format({'border': 1, 'bg_color': cor_ama})
                    
                    f_vde = wb.add_format({'num_format': '#,##0.00', 'font_color': 'green', 'bold': 1, 'border': 1})
                    f_vrm = wb.add_format({'num_format': '#,##0.00', 'font_color': 'red', 'bold': 1, 'border': 1})

                    for cod, df in banco.items():
                        ws = wb.add_worksheet(str(cod)[:31])
                        ws.hide_gridlines(2)
                        ws.set_column('A:A', 2); ws.set_column('B:C', 15); ws.set_column('D:D', 45); ws.set_column('E:F', 18); ws.set_column('I:L', 18)
                        
                        ws.merge_range('B2:L2', f"EMPRESA: {nome_emp}", f_emp)
                        ws.merge_range('B4:F4', f_info[cod], f_cab)
                        ws.merge_range('I4:L4', "CONCILIAÇÃO POR NOTA", f_cab)
                        
                        # Tabela Esquerda
                        for ci, v in enumerate(["Data","NF","Histórico","Débito","Crédito"]): ws.write(5, ci+1, v, f_cab)
                        for ri, r in enumerate(df.values):
                            fmt_c = f_ama_c if r[5] else f_c
                            fmt_m = f_ama_m if r[5] else f_m
                            fmt_s = f_ama_s if r[5] else f_s
                            
                            ws.write(6+ri, 1, r[0], fmt_c)
                            
                            # Escrevendo a NF como número se possível para tirar o cantinho verde
                            try:
                                if r[1].isdigit(): ws.write_number(6+ri, 2, int(r[1]), fmt_c)
                                else: ws.write(6+ri, 2, r[1], fmt_c)
                            except: ws.write(6+ri, 2, r[1], fmt_c)
                            
                            ws.write(6+ri, 3, r[2], fmt_s)
                            ws.write_number(6+ri, 4, r[3], fmt_m)
                            ws.write_number(6+ri, 5, r[4], fmt_m)
                            row_f = 6+ri
                        
                        # Tabela Direita
                        res = df.groupby("NF").agg({"Deb":"sum", "Cred":"sum"}).reset_index()
                        res["Dif"] = res["Deb"] + res["Cred"]
                        for ci, v in enumerate(["NF","Deb","Cred","Dif"]): ws.write(5, ci+8, v, f_cab)
                        for ri, r in enumerate(res.values):
                            try:
                                if str(r[0]).isdigit(): ws.write_number(6+ri, 8, int(r[0]), f_c)
                                else: ws.write(6+ri, 8, str(r[0]), f_c)
                            except: ws.write(6+ri, 8, str(r[0]), f_c)
                            
                            ws.write_number(6+ri, 9, r[1], f_m)
                            ws.write_number(6+ri, 10, r[2], f_m)
                            ws.write_number(6+ri, 11, r[3], f_m)
                            row_f_res = 6+ri
                        
                        sf_row = max(row_f, row_f_res) + 2
                        ws.write(sf_row, 10, "Saldo Final:", f_cab)
                        s = res["Dif"].sum()
                        ws.write_number(sf_row, 11, s, f_vde if abs(s) < 0.01 else f_vrm)

                st.success("✅ Relatório pronto! Sem avisos verdes e com o layout SOLUX restaurado.")
                st.download_button("📥 Baixar Relatório", out.getvalue(), "relatorio_solux.xlsx")
        except Exception as e:
            st.error(f"Erro: {e}")
