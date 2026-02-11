import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="SOLUX 2026", layout="wide")

st.title("💡 SOLUX 2026 - Gerador de Arquivo")

def to_num(val):
    try:
        s = str(val).replace('.', '').replace(',', '.')
        return float(re.sub(r'[^-0-9.]', '', s))
    except: return 0.0

tipo_robo = st.sidebar.radio("Tipo:", ["Cliente", "Fornecedor"])
arquivo = st.sidebar.file_uploader("Suba o arquivo aqui", type=["xlsx", "xls", "csv"])

if arquivo:
    try:
        # Lendo o arquivo
        if arquivo.name.endswith('.csv'):
            df = pd.read_csv(arquivo, header=None, encoding='latin-1')
        else:
            df = pd.read_excel(arquivo, header=None)

        # Processando as linhas
        dados = []
        for i in range(len(df)):
            l = df.iloc[i]
            # Procura linhas que tenham data (ex: 01/01/2026)
            if len(l) > 9 and pd.notna(l[0]) and '/' in str(l[0]):
                d, c = to_num(l[8]), to_num(l[9])
                h = str(l[2])
                
                # Busca nota (SAÍDA, PRESTADO, NF...)
                pats = [r'SAÍDA\s?(\d+)', r'PRESTADO\s?(\d+)', r'NF\s?(\d+)', r'NFE\s?(\d+)']
                nf = "S/N"
                for p in pats:
                    m = re.findall(p, h.upper())
                    if m: nf = m[0]; break
                
                # Regra de sinais do contexto
                v_d, v_c = (-d, c) if tipo_robo == "Fornecedor" else (d, -c)
                dados.append({"Data": l[0], "Historico": h, "NF_AJUSTADA": nf, "Debito": v_d, "Credito": v_c})

        if dados:
            df_razao = pd.DataFrame(dados)
            # Criando o Resumo por Nota (A "Dinâmica" pronta)
            df_resumo = df_razao.groupby("NF_AJUSTADA").agg({"Debito":"sum", "Credito":"sum"}).reset_index()
            df_resumo["Diferença"] = df_resumo["Debito"] + df_resumo["Credito"]

            # Criando o Excel na memória
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_razao.to_excel(writer, sheet_name='1-RAZAO', index=False)
                df_resumo.to_excel(writer, sheet_name='2-CONCILIACAO', index=False)

            # MOSTRANDO O BOTÃO
            st.success("Prontinho! O botão apareceu abaixo:")
            st.download_button(
                label="📥 BAIXAR EXCEL AGORA",
                data=output.getvalue(),
                file_name="solux_resultado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Não achei dados no arquivo. Verifique se é o modelo correto.")

    except Exception as e:
        st.error(f"Ocorreu um erro: {e}")
