import streamlit as st
import pandas as pd
import re
from io import BytesIO

# Configuração da Página
st.set_page_config(page_title="SOLUX 2026", page_icon="💡", layout="wide")

# Estilo SOLUX Lilás
st.markdown("""
    <style>
    .titulo { font-family: 'sans-serif'; color: #4B0082; font-size: 24px; font-weight: bold; text-align: center; padding: 15px; background-color: #E6E0FF; border-radius: 10px; }
    .stDownloadButton button { background-color: #9B8ADE !important; color: white !important; width: 100%; height: 50px; font-weight: bold; }
    </style>
    <div class="titulo">💡 SOLUX 2026: Conciliação Garantida 💡</div>
    """, unsafe_allow_html=True)

def to_num(val):
    try:
        if pd.isna(val) or str(val).strip() == '': return 0.0
        s = str(val).replace('.', '').replace(',', '.')
        return float(re.sub(r'[^-0-9.]', '', s))
    except: return 0.0

with st.sidebar:
    st.header("⚙️ Painel")
    tipo_robo = st.radio("Projeto de:", ["Cliente", "Fornecedor"])
    arquivo = st.file_uploader("Suba o arquivo aqui", type=["xlsx", "xls", "csv"])

if arquivo:
    try:
        # Lendo os dados
        if arquivo.name.endswith('.csv'):
            df_bruto = pd.read_csv(arquivo, header=None, engine='python', encoding='latin-1')
        else:
            df_bruto = pd.read_excel(arquivo, header=None)

        dados_lista = []
        for i in range(len(df_bruto)):
            lin = df_bruto.iloc[i]
            # Procura as linhas de lançamento (que tem data)
            if len(lin) >= 10 and pd.notna(lin[0]) and '/' in str(lin[0]):
                deb, cre = to_num(lin[8]), to_num(lin[9])
                if deb != 0 or cre != 0:
                    hist = str(lin[2]).strip()
                    h_up = hist.upper()
                    
                    # Busca NF: SAÍDA, PRESTADO, NF, etc.
                    pats = [r'SAÍDA\s?(\d+)', r'PRESTADO\s?(\d+)', r'NF\s?(\d+)', r'NFE\s?(\d+)', r'CTE\s?(\d+)']
                    nf_res = "S/N"
                    for p in pats:
                        m = re.findall(p, h_up)
                        if m: nf_res = m[0]; break
                    
                    # Regras de sinais do seu contexto (Saved Information)
                    v_deb, v_cre = (-deb, cre) if tipo_robo == "Fornecedor" else (deb, -cre)
                    dados_lista.append({"Data": lin[0], "Historico": hist, "NF_AJUSTADA": nf_res, "Debito": v_deb, "Credito": v_cre})

        if dados_lista:
            df_final = pd.DataFrame(dados_lista)
            
            # CRIANDO A CONCILIAÇÃO (A "Dinâmica" que já vem pronta)
            df_resumo = df_final.groupby("NF_AJUSTADA").agg({"Debito":"sum", "Credito":"sum"}).reset_index()
            df_resumo["DIFERENÇA"] = df_resumo["Debito"] + df_resumo["Credito"]
            df_resumo["STATUS"] = df_resumo["DIFERENÇA"].apply(lambda x: "OK" if abs(x) < 0.01 else "ABERTO")

            # Gerando o Excel com 2 abas
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Aba 1: O Razão com a coluna NF_AJUSTADA para você conferir
                df_final.to_excel(writer, sheet_name='1. RAZÃO DADOS', index=False)
                # Aba 2: A Conciliação prontinha (O que você chamou de Dinâmica)
                df_resumo.to_excel(writer, sheet_name='2. CONCILIAÇÃO PRONTA', index=False)
                
                # Deixando as colunas bonitas
                for aba in writer.sheets:
                    writer.sheets[aba].set_column('A:F', 20)

            # AGORA O BOTÃO APARECE!
            st.success("✅ Consegui! A conciliação está pronta no botão abaixo:")
            st.download_button(
                label="📥 BAIXAR CONCILIAÇÃO SOLUX",
                data=output.getvalue(),
                file_name=f"conciliacao_{tipo_robo.lower()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ Não encontrei lançamentos com valores no seu arquivo.")

    except Exception as e:
        st.error(f"❌ Ops! Ocorreu este erro: {e}")
