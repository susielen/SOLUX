import streamlit as st
import pandas as pd
import re
from io import BytesIO

# 1. Configuração da Página
st.set_page_config(page_title="SOLUX 2026", page_icon="💡", layout="wide")

# Estilo SOLUX (Lilás e Branco)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;800&display=swap');
    .stApp { background-color: #F3F0FF; }
    .titulo { font-family: 'Montserrat', sans-serif; color: #4B0082; font-size: 28px; font-weight: 800; text-align: center; padding: 10px; background-color: rgba(230, 224, 255, 0.9); border-radius: 10px; border: 1px solid #9B8ADE; margin-bottom: 25px; }
    .stDownloadButton button { background-color: #9B8ADE !important; color: white !important; border-radius: 8px !important; width: 100%; height: 50px; font-size: 20px; font-weight: bold; }
    </style>
    <p class="titulo">💡 SOLUX 2026: Conciliação Inteligente 💡</p>
    """, unsafe_allow_html=True)

def to_num(val):
    try:
        if pd.isna(val) or str(val).strip() == '': return 0.0
        s = str(val).replace('.', '').replace(',', '.')
        return float(re.sub(r'[^-0-9.]', '', s))
    except: return 0.0

with st.sidebar:
    st.header("⚙️ Painel de Controle")
    tipo_robo = st.radio("Este projeto é de:", ["Cliente", "Fornecedor"])
    arquivo = st.file_uploader("Suba o arquivo aqui", type=["xlsx", "xls", "csv"])

if arquivo:
    # Criamos um espaço vazio para colocar o botão depois
    placeholder = st.empty()
    
    with st.spinner('O robô SOLUX está processando... 🕵️‍♂️✨'):
        try:
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
                        # Lista de busca: NF, SAÍDA, PRESTADO, etc.
                        pats = [r'SERVIÇO\s?PRESTADO\s?(\d+)', r'NF\s?DE\s?S\s?(\d+)', r'FRETE\s?TOMADO\s?(\d+)', r'CTE\s?(\d+)', r'NFE\s?(\d+)', r'SAÍDA\s?(\d+)', r'NF\s?(\d+)']
                        nf_res = "S/N"
                        for p in pats:
                            m = re.findall(p, h_up)
                            if m: nf_res = m[0]; break
                        
                        # Regra de Sinais (Contexto do Usuário)
                        v_deb, v_cre = (-deb, cre) if tipo_robo == "Fornecedor" else (deb, -cre)
                        dados_lista.append({"Data": str(lin[0]), "Historico": hist, "NF_AJUSTADA": nf_res, "Debito": v_deb, "Credito": v_cre})

            if dados_lista:
                df_final = pd.DataFrame(dados_lista)
                out = BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    # ABA 1: RAZÃO
                    df_final.to_excel(writer, sheet_name='Razao', index=False)
                    # ABA 2: CONCILIAÇÃO
                    df_resumo = df_final.groupby("NF_AJUSTADA").agg({"Debito":"sum", "Credito":"sum"}).reset_index()
                    df_resumo["Diferença"] = df_resumo["Debito"] + df_resumo["Credito"]
                    df_resumo.to_excel(writer, sheet_name='Conciliacao', index=False)
                
                # SUCESSO E BOTÃO
                st.success("✅ Arquivo processado com sucesso!")
                st.download_button(
                    label="📥 CLIQUE AQUI PARA BAIXAR O ARQUIVO",
                    data=out.getvalue(),
                    file_name=f"conciliacao_{tipo_robo.lower()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Não encontrei dados para processar dentro desse arquivo.")

        except Exception as e:
            st.error(f"Erro ao processar: {e}")
