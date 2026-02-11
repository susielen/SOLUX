import streamlit as st
import pandas as pd
import re
from io import BytesIO

# 1. Configuração da Página
st.set_page_config(
    page_title="SOLUX",
    page_icon="💡",
    layout="wide"
)

# 2. O ESTILO (Lavanda Suave, Título Fino e Botão Camaleão)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;800&display=swap');

    .stApp {
        background-color: #F3F0FF; 
        background-image: url("https://www.transparenttextures.com/patterns/cubes.png");
        background-attachment: fixed;
    }

    header[data-testid="stHeader"], [data-testid="stSidebar"] {
        background-color: #9B8ADE !important;
    }

    button[kind="headerNoPadding"], .stApp header svg {
        display: none !important;
    }

    .titulo {
        font-family: 'Montserrat', sans-serif;
        color: #4B0082;
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

    [data-testid="stSidebar"] * {
        font-family: 'Montserrat', sans-serif;
        color: #FFFFFF !important;
        font-weight: 600 !important;
    }

    /* BROWSE FILES CAMALEÃO */
    [data-testid="stFileUploaderDropzone"] button {
        background-color: #E6E0FF !important;
        color: #4B0082 !important;
        border: 1px solid #9B8ADE !important;
        transition: 0.3s;
    }

    [data-testid="stFileUploaderDropzone"] button:hover {
        background-color: #9B8ADE !important;
        color: white !important;
    }

    [data-testid="stFileUploaderDropzone"] {
        background-color: rgba(255, 255, 255, 0.4) !important; 
        border: 2px dashed #9B8ADE !important;
    }
    
    .stDownloadButton button {
        background-color: #4B0082 !important;
        color: white !important;
        border-radius: 8px !important;
    }
    </style>
    
    <p class="titulo">💡 SOLUX: Seu parceiro na conciliação 💡</p>
    """, unsafe_allow_html=True)

# 3. FUNÇÃO DE NÚMEROS
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
    with st.spinner('💎 O SOLUX está lapidando os dados para sua Tabela Dinâmica...'):
        try:
            if arquivo.name.endswith('.csv'):
                df_bruto = pd.read_csv(arquivo, header=None, sep=None, engine='python', encoding='latin-1')
            else:
                df_bruto = pd.read_excel(arquivo, header=None)

            # Localizar Nome da Empresa
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
                        
                        # BUSCA DE NF NO HISTÓRICO (Se você editou no Excel, a Dinâmica lerá)
                        busca_nf = re.findall(r'(?:NF|NFE|NOTA|Nf|nfe|Nº|N)\s?(\d+)', hist)
                        nf = busca_nf[0] if busca_nf else (str(lin[1]).strip() if pd.notna(lin[1]) else "S/N")
                        
                        # REGRA DE SINAIS CONFORME SUA PREFERÊNCIA
                        if tipo_robo == "Fornecedor": val_deb, val_cre = -deb, cre
                        else: val_deb, val_cre = deb, -cre
                        
                        dados.append({"Data": str(lin[0]), "NF": nf, "Hist": hist, "Deb": val_deb, "Cred": val_cre})

            if f_cod and dados: banco[f_cod] = pd.DataFrame(dados)

            if banco:
                out = BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    for cod, df_res in banco.items():
                        sheet_name = str(cod)[:31]
                        df_res.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        # AQUI ESTÁ O SEGREDO: CRIAR A TABELA OFICIAL NO EXCEL
                        worksheet = writer.sheets[sheet_name]
                        (max_row, max_col) = df_res.shape
                        column_settings = [{'header': column} for column in df_res.columns]
                        # Cria a tabela que a Dinâmica adora
                        worksheet.add_table(0, 0, max_row, max_col - 1, {
                            'columns': column_settings,
                            'style': 'TableStyleLight 9'
                        })
                
                st.success("✅ Excel gerado! Agora use a Tabela Dinâmica para conciliar.")
                st.download_button("📥 BAIXAR EXCEL DINÂMICO 💎", out.getvalue(), "solux_conciliacao.xlsx")

        except Exception as e:
            st.error(f"Erro: {e}")
else:
    st.info("👋 O SOLUX está esperando o arquivo para organizar suas tabelas!")
