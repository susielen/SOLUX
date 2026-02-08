import streamlit as st
import pandas as pd
import re
from io import BytesIO
import time

# 1. Configuração da Página
st.set_page_config(
    page_title="SOLUX",
    page_icon="💡",
    layout="wide"
)

# 2. O ESTILO (Visual Solux Lilás)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;800&display=swap');
    .stApp { background-color: #F3F0FF; background-image: url("https://www.transparenttextures.com/patterns/cubes.png"); background-attachment: fixed; }
    header[data-testid="stHeader"], [data-testid="stSidebar"] { background-color: #9B8ADE !important; }
    .titulo { font-family: 'Montserrat', sans-serif; color: #4B0082; font-size: 28px; font-weight: 800; text-align: center; padding: 10px; background-color: rgba(230, 224, 255, 0.9); border-radius: 10px; border: 1px solid #9B8ADE; margin-top: -35px; margin-bottom: 25px; }
    [data-testid="stSidebar"] * { font-family: 'Montserrat', sans-serif; color: #FFFFFF !important; font-weight: 600 !important; }
    .stDownloadButton button { background-color: #9B8ADE !important; color: white !important; border-radius: 8px !important; border: none !important; }
    </style>
    <p class="titulo">💡 SOLUX: Conciliação Inteligente 💡</p>
    """, unsafe_allow_html=True)

# 3. FUNÇÃO PARA NÚMEROS
def to_num(val):
    try:
        if pd.isna(val) or str(val).strip() == '': return 0.0
        return float(str(val).replace('.', '').replace(',', '.'))
    except: return 0.0

# 4. PAINEL LATERAL
with st.sidebar:
    st.header("⚙️ Painel de Controle")
    tipo_robo = st.radio("Este projeto é de:", ["Cliente", "Fornecedor"])
    st.markdown("---")
    arquivo = st.file_uploader("Suba o arquivo aqui", type=["xlsx", "xls", "csv"])

# 5. LÓGICA DE CONCILIAÇÃO
if arquivo:
    with st.spinner('O robô está separando as notas... 🕵️‍♂️'):
        df_bruto = None
        try:
            if arquivo.name.endswith('.csv'):
                df_bruto = pd.read_csv(arquivo, header=None, sep=None, engine='python', encoding='latin-1')
            elif arquivo.name.endswith('.xls'):
                df_bruto = pd.read_excel(arquivo, header=None, engine='xlrd')
            else:
                df_bruto = pd.read_excel(arquivo, header=None)
        except Exception as e:
            st.error(f"⚠️ Erro ao abrir: {e}")

        if df_bruto is not None:
            try:
                nome_emp = "EMPRESA"
                for i in range(min(15, len(df_bruto))):
                    if "Empresa:" in str(df_bruto.iloc[i, 0]):
                        nome_emp = str(df_bruto.iloc[i, 2]); break

                banco, f_info = {}, {}
                f_cod, dados = [] , []

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
                            
                            # --- LÓGICA DE BUSCA SOLUX (SÓ PEGA SE TIVER TERMO CHAVE) ---
                            h_up = hist.upper()
                            
                            termos_chave = [
                                r'SERVIÇO\s?PRESTADO\s?(\d+)', 
                                r'NF\s?DE\s?S\s?(\d+)',         
                                r'FRETE\s?TOMADO\s?(\d+)',      
                                r'CTE\s?(\d+)',                 
                                r'NFE\s?(\d+)',                 
                                r'SAÍDA\s?(\d+)',
                                r'NF\s?(\d+)'
                            ]
                            
                            nf_identificada = None
                            for t in termos_chave:
                                achou = re.findall(t, h_up)
                                if achou:
                                    nf_identificada = achou[0]
                                    break
                            
                            # REGRA FINAL: Se não tem termo chave no histórico, é "S/ N° NF"
                            # Ignoramos até a coluna de documento se não houver prova no histórico
                            nf = nf_identificada if nf_identificada else "S/ N° NF"
                            
                            # SINAIS: Cliente (Débito + / Crédito -) | Fornecedor (Débito - / Crédito +)
                            if tipo_robo == "Fornecedor": val_deb, val_cre = -deb, cre
                            else: val_deb, val_cre = deb, -cre
                            
                            dados.append({"Data": dt, "NF": nf, "Hist": hist, "Deb": val_deb, "Cred": val_cre})

                if f_cod and dados: banco[f_cod] = pd.DataFrame(dados)

                if banco:
                    out = BytesIO()
                    with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                        wb = writer.book
                        f_cent = wb.add_format({'align': 'center', 'border': 1})
                        f_moeda = wb.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
                        f_std = wb.add_format({'border': 1})
                        f_cab = wb.add_format({'bold': 1, 'bg_color': '#F2F2F2', 'align': 'center', 'border': 1})
                        f_emp = wb.add_format({'bold': 1, 'font_size': 14, 'align': 'center', 'bg_color': '#D3D3D3', 'border': 1})
                        f_vde = wb.add_format({'num_format': 'R$ #,##0.00', 'font_color': 'green', 'bold': 1, 'border': 1})
                        f_vrm = wb.add_format({'num_format': 'R$ #,##0.00', 'font_color': 'red', 'bold': 1, 'border': 1})

                        for cod, df in banco.items():
                            ws = wb.add_worksheet(str(cod)[:31])
                            ws.hide_gridlines(2)
                            ws.set_column('B:C', 15); ws.set_column('D:D', 45); ws.set_column('E:F', 18); ws.set_column('I:L', 18)
                            ws.merge_range('B2:L2', f"EMPRESA: {nome_emp}", f_emp)
                            ws.merge_range('B4:F4', f_info[cod], f_cab)
                            ws.merge_range('I4:L4', "CONCILIAÇÃO POR NOTA", f_cab)
                            
                            for ci, v in enumerate(["Data","NF","Histórico","Débito","Crédito"]):
                                ws.write(5, ci+1, v, f_cab)
                            
                            for ri, row in enumerate(df.values):
                                ws.write(6+ri, 1, row[0], f_cent); ws.write(6+ri, 2, row[1], f_cent)
                                ws.write(6+ri, 3, row[2], f_std); ws.write(6+ri, 4, row[3], f_moeda); ws.write(6+ri, 5, row[4], f_moeda)
                            
                            # Tabela de Conciliação
                            res = df.groupby("NF").agg({"Deb":"sum","Cred":"sum"}).reset_index()
                            res["Dif"] = res["Deb"] + res["Cred"]
                            for ci, v in enumerate(["NF","Deb","Cred","Dif"]): ws.write(5, ci+8, v, f_cab)
                            for ri, row in enumerate(res.values):
                                ws.write(6+ri, 8, str(row[0]), f_cent)
                                ws.write(6+ri, 9, row[1], f_m := f_moeda); ws.write(6+ri, 10, row[2], f_m); ws.write(6+ri, 11, row[3], f_m)
                            
                            rf = 7 + len(res)
                            ws.write(rf, 10, "Saldo Final:", f_cab)
                            s = res["Dif"].sum()
                            ws.write(rf, 11, s, f_vde if abs(s) < 0.01 else f_vrm)

                    st.success("✅ Relatório gerado! Tudo o que não tinha termo de nota foi para 'S/ N° NF'.")
                    st.download_button("📥 Baixar Relatório SOLUX", out.getvalue(), "relatorio_final.xlsx")
            except Exception as e:
                st.error(f"⚠️ Erro no processamento: {e}")
