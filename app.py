import streamlit as st
import pandas as pd
import re
from io import BytesIO

# 1. Configuração da Página
st.set_page_config(page_title="SOLUX 2026", page_icon="💡", layout="wide")

# Estilo SOLUX Lilás
st.markdown("""
    <style>
    .titulo { font-family: 'sans-serif'; color: #4B0082; font-size: 26px; font-weight: bold; text-align: center; padding: 15px; background-color: #E6E0FF; border-radius: 10px; margin-bottom: 20px; }
    .stDownloadButton button { background-color: #9B8ADE !important; color: white !important; width: 100%; height: 60px; font-weight: bold; font-size: 20px; }
    </style>
    <div class="titulo">💡 SOLUX 2026: Layout Especial Lado a Lado 💡</div>
    """, unsafe_allow_html=True)

def to_num(val):
    try:
        if pd.isna(val) or str(val).strip() == '': return 0.0
        s = str(val).replace('.', '').replace(',', '.')
        return float(re.sub(r'[^-0-9.]', '', s))
    except: return 0.0

with st.sidebar:
    st.header("⚙️ Configurações")
    tipo_robo = st.radio("Projeto de:", ["Cliente", "Fornecedor"])
    arquivo = st.file_uploader("Suba o arquivo aqui", type=["xlsx", "xls", "csv"])

if arquivo:
    try:
        # Lendo os dados
        if arquivo.name.endswith('.csv'):
            df_bruto = pd.read_csv(arquivo, header=None, engine='python', encoding='latin-1')
        else:
            df_bruto = pd.read_excel(arquivo, header=None)

        # Identificando o nome da empresa e da conta
        nome_empresa = "EMPRESA NÃO IDENTIFICADA"
        nome_conta = "CONTA NÃO IDENTIFICADA"
        
        for i in range(min(20, len(df_bruto))):
            celula = str(df_bruto.iloc[i, 0])
            if "EMPRESA:" in celula.upper():
                nome_empresa = str(df_bruto.iloc[i, 0])
            if "CONTA:" in celula.upper() or "150 -" in celula:
                nome_conta = str(df_bruto.iloc[i, 0])

        dados_lista = []
        for i in range(len(df_bruto)):
            lin = df_bruto.iloc[i]
            if len(lin) >= 10 and pd.notna(lin[0]) and '/' in str(lin[0]):
                deb, cre = to_num(lin[8]), to_num(lin[9])
                if deb != 0 or cre != 0:
                    hist = str(lin[2]).strip()
                    h_up = hist.upper()
                    
                    pats = [r'SAÍDA\s?(\d+)', r'PRESTADO\s?(\d+)', r'NF\s?DE\s?S\s?(\d+)', r'NF\s?(\d+)', r'NFE\s?(\d+)']
                    nf_res = "S/N"
                    for p in pats:
                        m = re.findall(p, h_up)
                        if m: nf_res = m[0]; break
                    
                    v_deb, v_cre = (-deb, cre) if tipo_robo == "Fornecedor" else (deb, -cre)
                    dados_lista.append({"Data": lin[0], "NF": nf_res, "Histórico": hist, "Débito": v_deb, "Crédito": v_cre})

        if dados_lista:
            df_final = pd.DataFrame(dados_lista)
            
            # Criando o Resumo para o lado direito
            df_resumo = df_final.groupby("NF").agg({"Débito":"sum", "Crédito":"sum"}).reset_index()
            df_resumo["DIFERENÇA"] = df_resumo["Débito"] + df_resumo["Crédito"]
            df_resumo["STATUS"] = df_resumo["DIFERENÇA"].apply(lambda x: "OK" if abs(x) < 0.01 else "EM ABERTO")

            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                ws = workbook.add_worksheet('Conciliação')
                
                # Formatos
                f_titulo = workbook.add_format({'bold': True, 'font_size': 12, 'bg_color': '#D9EAD3', 'border': 1})
                f_cab = workbook.add_format({'bold': True, 'bg_color': '#EFEFEF', 'border': 1, 'align': 'center'})
                f_num = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
                f_txt = workbook.add_format({'border': 1})
                f_cnt = workbook.add_format({'border': 1, 'align': 'center'})

                # Escrevendo o Cabeçalho (Linha 1 e 2)
                ws.merge_range('A1:L1', nome_empresa, f_titulo)
                ws.merge_range('A2:E2', nome_conta, f_titulo)
                ws.merge_range('H2:L2', "CONCILIAÇÃO POR NOTA", f_titulo)

                # Cabeçalhos das colunas (Linha 4)
                colunas_razao = ["Data", "NF", "Histórico", "Débito", "Crédito"]
                for i, col in enumerate(colunas_razao):
                    ws.write(3, i, col, f_cab)
                
                colunas_resumo = ["NF", "Crédito", "Débito", "DIFERENÇA", "STATUS"]
                for i, col in enumerate(colunas_resumo):
                    ws.write(3, i + 7, col, f_cab)

                # Escrevendo Dados do Razão (Lado Esquerdo)
                for r_idx, row in df_final.iterrows():
                    ws.write(4 + r_idx, 0, str(row['Data']), f_cnt)
                    ws.write(4 + r_idx, 1, row['NF'], f_cnt)
                    ws.write(4 + r_idx, 2, row['Histórico'], f_txt)
                    ws.write(4 + r_idx, 3, row['Débito'], f_num)
                    ws.write(4 + r_idx, 4, row['Crédito'], f_num)

                # Escrevendo Dados da Dinâmica (Lado Direito)
                for r_idx, row in df_resumo.iterrows():
                    ws.write(4 + r_idx, 7, row['NF'], f_cnt)
                    ws.write(4 + r_idx, 8, row['Crédito'], f_num)
                    ws.write(4 + r_idx, 9, row['Débito'], f_num)
                    ws.write(4 + r_idx, 10, row['DIFERENÇA'], f_num)
                    ws.write(4 + r_idx, 11, row['STATUS'], f_cnt)

                # Ajuste de largura das colunas
                ws.set_column('A:B', 12); ws.set_column('C:C', 40); ws.set_column('D:E', 15)
                ws.set_column('G:G', 5); ws.set_column('H:L', 15)

            st.success("✅ Arquivo gerado exatamente no modelo solicitado!")
            st.download_button(
                label="📥 BAIXAR EXCEL SOLUX (MODELO LADO A LADO)",
                data=output.getvalue(),
                file_name=f"conciliacao_solux_oficial.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Não encontrei dados para processar.")

    except Exception as e:
        st.error(f"Erro: {e}")
