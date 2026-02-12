import streamlit as st
import pandas as pd
import re
from io import BytesIO

# 1. Configuração da Página
st.set_page_config(page_title="SOLUX", page_icon="💡", layout="wide")

# 2. ESTILO SOLUX QUE VOCÊ GOSTA (Cores, Fontes e Botão Lilás)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;800&display=swap');
    .stApp { background-color: #F3F0FF; background-image: url("https://www.transparenttextures.com/patterns/cubes.png"); background-attachment: fixed; }
    header[data-testid="stHeader"], [data-testid="stSidebar"] { background-color: #9B8ADE !important; }
    .titulo { font-family: 'Montserrat', sans-serif; color: #4B0082; font-size: 28px; font-weight: 800; text-align: center; padding: 10px; background-color: rgba(230, 224, 255, 0.9); border-radius: 10px; border: 1px solid #9B8ADE; margin-top: -35px; margin-bottom: 25px; }
    
    [data-testid="stFileUploaderDropzone"] {
        background-color: rgba(255, 255, 255, 0.6) !important;
        border: 2px dashed #9B8ADE !important;
    }
    
    [data-testid="stSidebar"] * { color: #FFFFFF !important; font-weight: 600 !important; }
    .stDownloadButton button { 
        background-color: #9B8ADE !important; 
        color: white !important; 
        border-radius: 8px !important; 
        width: 100%; 
        height: 60px; 
        font-size: 20px; 
        font-weight: bold;
        box-shadow: 0px 4px 10px rgba(0,0,0,0.1);
    }
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
    with st.spinner('SOLUX está preparando o seu layout favorito... 🕵️‍♂️✨'):
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
                        # Busca termos específicos conforme sua instrução [2026-02-05]
                        pats = [r'SAÍDA\s?(\d+)', r'PRESTADO\s?(\d+)', r'NF\s?DE\s?S\s?(\d+)', r'CTE\s?(\d+)', r'NFE\s?(\d+)', r'NF\s?(\d+)']
                        nf_res = None
                        for p in pats:
                            m = re.findall(p, h_up)
                            if m: nf_res = m[0]; break
                        
                        nf = nf_res if nf_res else "S/ N° NF"
                        
                        # Regra de sinais conforme sua instrução [2026-01-30]
                        if tipo_robo == "Fornecedores":
                            v_deb, v_cre = -deb, cre
                        else:
                            v_deb, v_cre = deb, -cre
                            
                        dados.append({"Data": data_formatada, "NF": nf, "Hist": hist, "Deb": v_deb, "Cred": v_cre, "Aviso": (nf == "S/ N° NF")})

            if f_cod and dados: banco[f_cod] = pd.DataFrame(dados)

            if banco:
                out = BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    wb = writer.book
                    # --- ESTILOS DO JEITO QUE VOCÊ GOSTA ---
                    f_cab = wb.add_format({'bold': 1, 'bg_color': '#9B8ADE', 'font_color': 'white', 'align': 'center', 'border': 1})
                    f_emp = wb.add_format({'bold': 1, 'font_size': 14, 'align': 'center', 'bg_color': '#E6E0FF', 'font_color': '#4B0082', 'border': 1})
                    f_c = wb.add_format({'align': 'center', 'border': 1})
                    f_m = wb.add_format({'num_format': '#,##0.00', 'border': 1})
                    f_s = wb.add_format({'border': 1})
                    f_vde = wb.add_format({'num_format': '#,##0.00', 'font_color': 'green', 'bold': 1, 'border': 1})
                    f_vrm = wb.add_format({'num_format': '#,##0.00', 'font_color': 'red', 'bold': 1, 'border': 1})
                    f_ama_c = wb.add_format({'align': 'center', 'border': 1, 'bg_color': '#FFFF99'}) # Amarelo para NF não encontrada

                    for cod, df_emp in banco.items():
                        ws = wb.add_worksheet(str(cod)[:31])
                        ws.hide_gridlines(2)
                        
                        ws.set_column('B:C', 15); ws.set_column('D:D', 45); ws.set_column('E:F', 18)
                        ws.set_column('G:H', 2.14) # Espaço entre as tabelas
                        ws.set_column('I:M', 18)
                        
                        ws.merge_range('B2:M2', f"RELATÓRIO SOLUX - {nome_emp}", f_emp)
                        ws.merge_range('B4:F4', f_info[cod], f_cab)
                        ws.merge_range('I4:M4', "RESUMO DINÂMICO", f_cab)

                        # --- LADO ESQUERDO: RAZÃO ---
                        for ci, v in enumerate(["Data","NF","Histórico","Débito","Crédito"]):
                            ws.write(5, ci+1, v, f_cab)
                        
                        row_f = 5
                        for ri, r in enumerate(df_emp.values):
                            fmt_c = f_ama_c if r[5] else f_c
                            ws.write(6+ri, 1, r[0], fmt_c)
                            ws.write(6+ri, 2, r[1], fmt_c)
                            ws.write(6+ri, 3, r[2], f_s)
                            ws.write_number(6+ri, 4, r[3], f_m)
                            ws.write_number(6+ri, 5, r[4], f_m)
                            row_f = 6+ri
                        
                        # Totais do Razão (com o pulo de linha que você gosta)
                        ws.write(row_f + 2, 3, "SOMA DOS TOTAIS:", f_cab)
                        ws.write_number(row_f + 2, 4, df_emp["Deb"].sum(), f_m)
                        ws.write_number(row_f + 2, 5, df_emp["Cred"].sum(), f_m)
                        
                        ws.write(row_f + 3, 4, "SALDO LÍQUIDO:", f_cab)
                        s_r = df_emp["Deb"].sum() + df_emp["Cred"].sum()
                        ws.write_number(row_f + 3, 5, s_r, f_vde if abs(s_r) < 0.01 else f_vrm)

                        # --- LADO DIREITO: DINÂMICA ---
                        res = df_emp.groupby("NF").agg({"Deb":"sum", "Cred":"sum"}).reset_index()
                        res["Dif"] = res["Deb"] + res["Cred"]
                        for ci, v in enumerate(["NF","Deb","Cred","DIFERENÇA", "STATUS"]):
                            ws.write(5, ci+8, v, f_cab)
                        
                        row_res = 5
                        for ri, r in enumerate(res.values):
                            ws.write(6+ri, 8, str(r[0]), f_c)
                            ws.write_number(6+ri, 9, r[1], f_m)
                            ws.write_number(6+ri, 10, r[2], f_m)
                            ws.write_number(6+ri, 11, r[3], f_m)
                            ok = abs(r[3]) < 0.01
                            ws.write(6+ri, 12, "CONCILIADO" if ok else "PENDENTE", 
                                     wb.add_format({'bold':1, 'border':1, 'align':'center', 'font_color': 'green' if ok else '#CC7A00'}))
                            row_res = 6+ri

                        # Saldo Final da Dinâmica
                        ws.write(row_res + 2, 11, "SALDO FINAL:", f_cab)
                        s_c = res["Dif"].sum()
                        ws.write_number(row_res + 2, 12, s_c, f_vde if abs(s_c) < 0.01 else f_vrm)

                st.success("✅ SOLUX! Conciliado com sucesso 😁.")
                st.download_button("📥 BAIXAR RELATÓRIO SOLUX FAVORITO", out.getvalue(), f"solux_{tipo_robo.lower()}_layout.xlsx")
        except Exception as e:
            st.error(f"Erro no processamento: {e}")
