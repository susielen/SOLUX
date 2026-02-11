import streamlit as st
import pandas as pd
import re
from io import BytesIO

# 1. Configuração da Página
st.set_page_config(page_title="SOLUX 2026", page_icon="💡", layout="wide")

# Estilo mantido
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;800&display=swap');
    .stApp { background-color: #F3F0FF; }
    .titulo { font-family: 'Montserrat', sans-serif; color: #4B0082; font-size: 28px; font-weight: 800; text-align: center; padding: 10px; background-color: rgba(230, 224, 255, 0.9); border-radius: 10px; border: 1px solid #9B8ADE; margin-bottom: 25px; }
    .stDownloadButton button { background-color: #9B8ADE !important; color: white !important; }
    </style>
    <p class="titulo">💡 SOLUX 2026: Tabela Dinâmica Automática 💡</p>
    """, unsafe_allow_html=True)

def to_num(val):
    try:
        if pd.isna(val) or str(val).strip() == '': return 0.0
        s = str(val).replace('.', '').replace(',', '.')
        return float(re.sub(r'[^-0-9.]', '', s))
    except: return 0.0

with st.sidebar:
    st.header("⚙️ Painel")
    tipo_robo = st.radio("Projeto:", ["Cliente", "Fornecedor"])
    arquivo = st.file_uploader("Arquivo", type=["xlsx", "xls", "csv"])

if arquivo:
    with st.spinner('Construindo a Tabela Dinâmica... 🏗️'):
        try:
            if arquivo.name.endswith('.csv'):
                df_bruto = pd.read_csv(arquivo, header=None, sep=None, engine='python', encoding='latin-1')
            else:
                df_bruto = pd.read_excel(arquivo, header=None)

            banco = {}
            f_cod, dados = None, []
            for i in range(len(df_bruto)):
                lin = df_bruto.iloc[i]
                if "Conta:" in str(lin[0]):
                    if f_cod and dados: banco[f_cod] = pd.DataFrame(dados)
                    f_cod = str(lin[1]).strip()
                    dados = []
                elif len(lin) >= 10 and pd.notna(lin[0]) and any(x in str(lin[0]) for x in ['/', '-']):
                    deb, cre = to_num(lin[8]), to_num(lin[9])
                    if deb != 0 or cre != 0:
                        hist = str(lin[2]).strip()
                        h_up = hist.upper()
                        pats = [r'SERVIÇO\s?PRESTADO\s?(\d+)', r'NF\s?DE\s?S\s?(\d+)', r'FRETE\s?TOMADO\s?(\d+)', r'CTE\s?(\d+)', r'NFE\s?(\d+)', r'SAÍDA\s?(\d+)', r'NF\s?(\d+)']
                        nf_res = "S/N"
                        for p in pats:
                            m = re.findall(p, h_up)
                            if m: nf_res = m[0]; break
                        
                        v_deb, v_cre = (-deb, cre) if tipo_robo == "Fornecedor" else (deb, -cre)
                        dados.append({"Data": str(lin[0]), "Hist": hist, "NF_AJUSTADA": nf_res, "Deb": v_deb, "Cred": v_cre})

            if f_cod and dados: banco[f_cod] = pd.DataFrame(dados)

            if banco:
                out = BytesIO()
                # Usando XlsxWriter para criar a estrutura de Pivot Table
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    for cod, df in banco.items():
                        aba = str(cod)[:31]
                        df.to_excel(writer, sheet_name=aba, index=False, startrow=5, startcol=1)
                        
                        wb = writer.book
                        ws = writer.sheets[aba]
                        
                        # 1. Transforma o Razão em uma TABELA OFICIAL (necessário para a Dinâmica)
                        # Removemos caracteres especiais do nome da tabela
                        nome_tab = f"Tabela_{re.sub(r'[^a-zA-Z0-9]', '', str(cod))}"
                        last_row = 5 + len(df)
                        ws.add_table(5, 1, last_row, 5, {
                            'name': nome_tab,
                            'columns': [{'header': 'Data'}, {'header': 'Hist'}, {'header': 'NF_AJUSTADA'}, {'header': 'Deb'}, {'header': 'Cred'}]
                        })

                        # 2. Instrução para o usuário (Já que Python não cria o "desenho" da pivot, mas sim a base)
                        ws.write('H4', 'COMO CRIAR A DINÂMICA:', wb.add_format({'bold': True, 'color': 'red'}))
                        ws.write('H5', '1. Clique em qualquer lugar da tabela azul à esquerda.')
                        ws.write('H6', '2. Vá em Inserir > Tabela Dinâmica.')
                        ws.write('H7', '3. Arraste "NF_AJUSTADA" para LINHAS e "Deb" e "Cred" para VALORES.')

                st.success("✅ Arquivo preparado com Tabela Nomeada!")
                st.download_button("📥 Baixar Relatório", out.getvalue(), "solux_dinamica_oficial.xlsx")

        except Exception as e:
            st.error(f"Erro: {e}")
