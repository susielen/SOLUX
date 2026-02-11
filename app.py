import streamlit as st
import pandas as pd
import re
from io import BytesIO

# 1. Configuração da Página
st.set_page_config(page_title="SOLUX 2026", page_icon="💡", layout="wide")

# Estilo SOLUX
st.markdown("""
    <style>
    .titulo { font-family: 'Montserrat', sans-serif; color: #4B0082; font-size: 28px; font-weight: 800; text-align: center; padding: 10px; background-color: rgba(230, 224, 255, 0.9); border-radius: 10px; border: 1px solid #9B8ADE; margin-bottom: 25px; }
    .stDownloadButton button { background-color: #9B8ADE !important; color: white !important; }
    </style>
    <p class="titulo">💡 SOLUX 2026: Tabela Dinâmica Pronta 💡</p>
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
    with st.spinner('O robô está montando a sua Tabela Dinâmica... 🏗️'):
        try:
            # Leitura dos dados
            if arquivo.name.endswith('.csv'):
                df_bruto = pd.read_csv(arquivo, header=None, sep=None, engine='python', encoding='latin-1')
            else:
                df_bruto = pd.read_excel(arquivo, header=None)

            # Processamento dos dados
            lista_final = []
            for i in range(len(df_bruto)):
                lin = df_bruto.iloc[i]
                if len(lin) >= 10 and pd.notna(lin[0]) and any(x in str(lin[0]) for x in ['/', '-']):
                    deb, cre = to_num(lin[8]), to_num(lin[9])
                    if deb != 0 or cre != 0:
                        hist = str(lin[2]).strip()
                        h_up = hist.upper()
                        # Busca de NF nas palavras-chave do contexto
                        pats = [r'SERVIÇO\s?PRESTADO\s?(\d+)', r'NF\s?DE\s?S\s?(\d+)', r'FRETE\s?TOMADO\s?(\d+)', r'CTE\s?(\d+)', r'NFE\s?(\d+)', r'SAÍDA\s?(\d+)', r'NF\s?(\d+)']
                        nf_res = "S/N"
                        for p in pats:
                            m = re.findall(p, h_up)
                            if m: nf_res = m[0]; break
                        
                        v_deb, v_cre = (-deb, cre) if tipo_robo == "Fornecedor" else (deb, -cre)
                        lista_final.append([str(lin[0]), hist, nf_res, v_deb, v_cre])

            df_final = pd.DataFrame(lista_final, columns=["Data", "Historico", "NF_AJUSTADA", "Debito", "Credito"])

            if not df_final.empty:
                out = BytesIO()
                # Criando o Excel com suporte a Tabelas Dinâmicas (XlsxWriter)
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    # 1. Aba de Dados (Razão)
                    df_final.to_excel(writer, sheet_name='Razao', index=False)
                    workbook = writer.book
                    worksheet_data = writer.sheets['Razao']
                    
                    # Criar Tabela Oficial para a Dinâmica ler
                    (max_row, max_col) = df_final.shape
                    column_settings = [{'header': column} for column in df_final.columns]
                    worksheet_data.add_table(0, 0, max_row, max_col - 1, {
                        'name': 'TabelaDados',
                        'columns': column_settings,
                        'style': 'TableStyleMedium 2'
                    })

                    # 2. Criar a Aba da Tabela Dinâmica
                    worksheet_pivot = workbook.add_worksheet('Conciliacao_Dinamica')
                    
                    # O XlsxWriter cria o "espaço" da Pivot. 
                    # Definimos NF_AJUSTADA em 'rows' e Debito/Credito em 'data'
                    worksheet_pivot.add_pivot_table(1, 1, 20, 5, 'TabelaDados', {
                        'rows': ['NF_AJUSTADA'],
                        'measures': [
                            {'function': 'sum', 'field': 'Debito', 'name': 'Soma de Débito'},
                            {'function': 'sum', 'field': 'Credito', 'name': 'Soma de Crédito'},
                        ],
                        'style': 'PivotStyleMedium 9'
                    })
                    
                    # Coluna de Diferença (Cálculo fora da pivot para garantir visibilidade)
                    worksheet_pivot.write('E2', 'Diferença Total', workbook.add_format({'bold': True, 'bg_color': '#9B8ADE', 'color': 'white'}))
                    worksheet_pivot.write_formula('E3', '=C3+D3')

                st.success("✅ Tabela Dinâmica criada com sucesso na segunda aba!")
                st.download_button("📥 Baixar Relatório SOLUX", out.getvalue(), "solux_com_dinamica.xlsx")

        except Exception as e:
            st.error(f"Erro no processamento: {e}")
