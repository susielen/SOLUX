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

# 2. O ESTILO (Cores Suaves, Letra Moderna e Fundo Profissional)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;800&display=swap');

    /* FUNDO LILÁS LAVANDA (Suave para os olhos) */
    .stApp {
        background-color: #F3F0FF; 
        background-image: url("https://www.transparenttextures.com/patterns/cubes.png");
        background-attachment: fixed;
    }

    /* BARRA LATERAL E TOPO (Roxo Pastel) */
    header[data-testid="stHeader"], [data-testid="stSidebar"] {
        background-color: #9B8ADE !important;
    }

    /* ESCONDER A COROA E ÍCONES DO TOPO */
    button[kind="headerNoPadding"], .stApp header svg {
        display: none !important;
    }

    /* TÍTULO FINO E ELEGANTE */
    .titulo {
        font-family: 'Montserrat', sans-serif;
        color: #4B0082; /* Roxo escuro para contraste */
        font-size: 28px; 
        font-weight: 800; 
        text-align: center; 
        padding: 8px; 
        background-color: rgba(230, 224, 255, 0.9);
        border-radius: 10px;
        border: 1px solid #9B8ADE;
        margin-top: -35px;
        margin-bottom: 25px;
    }

    /* TEXTOS DA BARRA LATERAL EM NEGRITO */
    [data-testid="stSidebar"] * {
        font-family: 'Montserrat', sans-serif;
        color: #FFFFFF !important;
        font-weight: 600 !important;
    }

    /* CAIXA DE UPLOAD LILÁS CLARO */
    [data-testid="stFileUploaderDropzone"] {
        background-color: rgba(255, 255, 255, 0.4) !important; 
        border: 2px dashed #9B8ADE !important;
        border-radius: 12px !important;
    }

    /* --- ESTA É A PARTE QUE MUDA A COR DO BROWSE FILES --- */
    [data-testid="stFileUploaderDropzone"] button {
        background-color: #E6E0FF !important; /* Cor igual a caixinha */
        color: #4B0082 !important; /* Letra roxa */
        border: 1px solid #9B8ADE !important;
        border-radius: 8px !important;
        transition: 0.3s; /* Deixa o efeito suave */
    }

    [data-testid="stFileUploaderDropzone"] button:hover {
        background-color: #9B8ADE !important; /* Cor igual ao corredor (lilás mais forte) */
        color: white !important; /* Letra fica branca */
    }
    
    /* BOTÃO DE DOWNLOAD PERSONALIZADO */
    .stDownloadButton button {
        background-color: #9B8ADE !important;
        color: white !important;
        border-radius: 8px !important;
        border: none !important;
        padding: 10px 20px !important;
    }
    </style>
    
    <p class="titulo">💡 SOLUX: Seu parceiro na conciliação 💡</p>
    """, unsafe_allow_html=True)

# 3. FUNÇÕES DE SUPORTE
def to_num(val):
    try:
        if pd.isna(val) or str(val).strip() == '': return 0.0
        return float(str(val).replace('.', '').replace(',', '.'))
    except: return 0.0

# 4. PAINEL LATERAL
with st.sidebar:
    st.header("⚙️ Painel de Controle")
    # Regra: Fornecedor (+/-) e Cliente (-/+) conforme conversamos
    tipo_robo = st.radio("Este projeto é de:", ["Cliente", "Fornecedor"])
    st.markdown("---")
    arquivo = st.file_uploader("Suba o arquivo aqui", type=["xlsx", "xls", "csv"])

# 5. LÓGICA DE CONCILIAÇÃO
if arquivo:
    with st.spinner('Conciliando...'):
        time.sleep(3)
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
if arquivo:
    with st.spinner('Conciliando...⏳⌛💱'):
        df_bruto = None
        try:
            if arquivo.name.endswith('.csv'):
                df_bruto = pd.read_csv(arquivo, header=None, sep=None, engine='python', encoding='latin-1')
            elif arquivo.name.endswith('.xls'):
                df_bruto = pd.read_excel(arquivo, header=None, engine='xlrd')
            else:
                df_bruto = pd.read_excel(arquivo, header=None)
        except Exception as e:
            st.error(f"⚠️ Erro ao abrir: {e}. Verifique o seu requirements.txt.")

        if df_bruto is not None:
            try:
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
                            nf_f = re.findall(r'NFe\s?(\d+)', hist)
                            nf = nf_f[0] if nf_f else (str(lin[1]).strip() if pd.notna(lin[1]) else "S/N")
                            
                            # REGRA DE SINAIS (Cliente: Débito + / Crédito -)
                            if tipo_robo == "Fornecedor": val_deb, val_cre = -deb, cre
                            else: val_deb, val_cre = deb, -cre
                            dados.append({"Data": dt, "NF": nf, "Hist": hist, "Deb": val_deb, "Cred": val_cre})

                if f_cod and dados: banco[f_cod] = pd.DataFrame(dados)

                if banco:
                    out = BytesIO()
                    with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                        wb = writer.book
                        
                        f_cent = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
                        f_moeda = wb.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
                        f_std = wb.add_format({'border': 1})
                        f_cab = wb.add_format({'bold': 1, 'bg_color': '#F2F2F2', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                        f_empresa = wb.add_format({'bold': 1, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D3D3D3', 'border': 1})
                        f_vde = wb.add_format({'num_format': 'R$ #,##0.00', 'font_color': 'green', 'bold': 1, 'border': 1, 'align': 'center'})
                        f_vrm = wb.add_format({'num_format': 'R$ #,##0.00', 'font_color': 'red', 'bold': 1, 'border': 1, 'align': 'center'})

                        for cod, df in banco.items():
                            ws = wb.add_worksheet(str(cod)[:31])
                            ws.hide_gridlines(2)
                            ws.ignore_errors({'number_stored_as_text': 'A1:X2000'})
                            
                            # LARGURA EXATA DE 2.0 PARA G e H
                            ws.set_column('A:A', 1); ws.set_column('B:C', 15); ws.set_column('D:D', 45)
                            ws.set_column('E:F', 18); ws.set_column('G:H', 2.0); ws.set_column('I:L', 18)
                            
                            ws.merge_range('B2:L2', f"EMPRESA: {nome_emp}", f_empresa)
                            ws.merge_range('B4:F4', f_info[cod], f_cab)
                            ws.merge_range('I4:L4', "CONCILIAÇÃO POR NOTA", f_cab)
                            
                            for ci, v in enumerate(["Data","NF","Histórico","Débito","Crédito"]):
                                ws.write(5, ci+1, v, f_cab)
                            
                            row_idx = 6
                            for ri, row in enumerate(df.values):
                                ws.write(row_idx+ri, 1, row[0], f_cent); ws.write(row_idx+ri, 2, row[1], f_cent)
                                ws.write(row_idx+ri, 3, row[2], f_std)
                                ws.write(row_idx+ri, 4, row[3], f_moeda); ws.write(row_idx+ri, 5, row[4], f_moeda)
                                last_row = row_idx + ri
                            
                            # TOTALIZADOR COM PULO DE UMA LINHA
                            lt = last_row + 2
                            ws.write(lt, 3, "TOTALIZADOR:", f_cab)
                            ws.write(lt, 4, df['Deb'].sum(), f_moeda)
                            ws.write(lt, 5, df['Cred'].sum(), f_moeda)
                            
                            # Conciliação
                            res = df.groupby("NF").agg({"Deb":"sum","Cred":"sum"}).reset_index()
                            res["Dif"] = res["Deb"] + res["Cred"]
                            for ci, v in enumerate(["NF","Deb","Cred","Dif"]): ws.write(5, ci+8, v, f_cab)
                            for ri, row in enumerate(res.values):
                                ws.write(6+ri, 8, str(row[0]), f_cent)
                                ws.write(6+ri, 9, row[1], f_moeda); ws.write(6+ri, 10, row[2], f_moeda); ws.write(6+ri, 11, row[3], f_moeda)
                            
                            rf = 7 + len(res)
                            ws.write(rf, 10, "Saldo Final:", f_cab)
                            ws.write(rf, 11, s := res["Dif"].sum(), f_vde if s >= 0 else f_vrm)
                    
                    st.success("✅ Relatório Lapidado com Sucesso!")
                    st.download_button("📥 Baixar Relatório ⬇️", out.getvalue(), "relatorio_lapidado.xlsx")
            except Exception as e:
                st.error(f"⚠️ Erro no processamento: {e}")
