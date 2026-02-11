import streamlit as st
import pandas as pd
import re
from io import BytesIO

# 1. Configuração da Página
st.set_page_config(page_title="SOLUX 2026", page_icon="💡", layout="wide")

# Estilo SOLUX (Lilás)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;800&display=swap');
    .stApp { background-color: #F3F0FF; }
    .titulo { font-family: 'Montserrat', sans-serif; color: #4B0082; font-size: 28px; font-weight: 800; text-align: center; padding: 10px; background-color: rgba(230, 224, 255, 0.9); border-radius: 10px; border: 1px solid #9B8ADE; margin-bottom: 25px; }
    .stDownloadButton button { background-color: #9B8ADE !important; color: white !important; border-radius: 8px !important; width: 100%; height: 60px; font-size: 22px; font-weight: bold; }
    </style>
    <p class="titulo">💡 SOLUX 2026: Conciliação com Tabela Dinâmica 💡</p>
    """, unsafe_allow_html=True)

def to_num(val):
    try:
        if pd.isna(val) or str(val).strip() == '': return 0.0
        s = str(val).replace('.', '').replace(',', '.')
        return float(re.sub(r'[^-0-9.]', '', s))
    except: return 0.0

with st.sidebar:
    st.header("⚙️ Configurações")
    tipo_robo = st.radio("Tipo de Projeto:", ["Cliente", "Fornecedor"])
    arquivo = st.file_uploader("Suba seu arquivo Excel/CSV", type=["xlsx", "xls", "csv"])

if arquivo:
    with st.spinner('Montando o Razão e a Tabela Dinâmica... 🏗️'):
        try:
            # Lendo os dados
            if arquivo.name.endswith('.csv'):
                df_bruto = pd.read_csv(arquivo, header=None, sep=None, engine='python', encoding='latin-1')
            else:
                df_bruto = pd.read_excel(arquivo, header=None)

            dados_lista = []
            for i in range(len(df_bruto)):
                lin = df_bruto.iloc[i]
                if len(lin) >= 10 and pd.notna(lin[0]) and any(x in str(lin[0]) for x in ['/', '-']):
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
                        dados_lista.append({"Data": str(lin[0]), "Historico": hist, "NF_AJUSTADA": nf_res, "Debito": v_deb, "Credito": v_cre})

            if dados_lista:
                df_final = pd.DataFrame(dados_lista)
                
                # Criando o arquivo Excel
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    # 1. ABA DO RAZÃO (DADOS)
                    df_final.to_excel(writer, sheet_name='Base_Dados', index=False)
                    workbook = writer.book
                    worksheet_data = writer.sheets['Base_Dados']
                    
                    # Criar a Tabela Oficial (Nomeada como 'TabelaRazao')
                    (max_row, max_col) = df_final.shape
                    column_settings = [{'header': column} for column in df_final.columns]
                    worksheet_data.add_table(0, 0, max_row, max_col - 1, {
                        'name': 'TabelaRazao',
                        'columns': column_settings,
                        'style': 'TableStyleMedium 2'
                    })

                    # 2. ABA DA CONCILIAÇÃO (DINÂMICA)
                    worksheet_pivot = workbook.add_worksheet('Conciliacao_Dinamica')
                    
                    # Montando a Tabela Dinâmica Automática
                    worksheet_pivot.add_pivot_table(1, 1, 20, 5, 'TabelaRazao', {
                        'rows': ['NF_AJUSTADA'],
                        'measures': [
                            {'function': 'sum', 'field': 'Debito', 'name': 'Soma de Debito'},
                            {'function': 'sum', 'field': 'Credito', 'name': 'Soma de Credito'},
                        ],
                        'style': 'PivotStyleMedium 9'
                    })
                    
                    # Espaço para o Status
                    worksheet_pivot.write('E1', 'Diferença Total', workbook.add_format({'bold': True}))
                    worksheet_pivot.write('F1', 'Status', workbook.add_format({'bold': True}))

                # MOSTRAR BOTÃO DE DOWNLOAD
                st.success("✨ O Razão e a Tabela Dinâmica estão prontos!")
                st.download_button(
                    label="📥 BAIXAR AGORA (VERSÃO COM DINÂMICA)",
                    data=output.getvalue(),
                    file_name=f"conciliacao_solux_{tipo_robo.lower()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Puxa, deu um erro aqui: {e}")
