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
    .stDownloadButton button { background-color: #9B8ADE !important; color: white !important; border-radius: 8px !important; }
    </style>
    <p class="titulo">💡 SOLUX 2026: Conciliação Automática 💡</p>
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
    with st.spinner('O robô SOLUX está reconstruindo tudo... 🕵️‍♂️✨'):
        try:
            if arquivo.name.endswith('.csv'):
                df_bruto = pd.read_csv(arquivo, header=None, sep=None, engine='python', encoding='latin-1')
            else:
                df_bruto = pd.read_excel(arquivo, header=None)

            # Processamento de Dados
            dados_lista = []
            for i in range(len(df_bruto)):
                lin = df_bruto.iloc[i]
                if len(lin) >= 10 and pd.notna(lin[0]) and any(x in str(lin[0]) for x in ['/', '-']):
                    deb, cre = to_num(lin[8]), to_num(lin[9])
                    if deb != 0 or cre != 0:
                        hist = str(lin[2]).strip()
                        h_up = hist.upper()
                        # Lista de busca conforme sua instrução
                        pats = [r'SERVIÇO\s?PRESTADO\s?(\d+)', r'NF\s?DE\s?S\s?(\d+)', r'FRETE\s?TOMADO\s?(\d+)', r'CTE\s?(\d+)', r'NFE\s?(\d+)', r'SAÍDA\s?(\d+)', r'NF\s?(\d+)']
                        nf_res = "S/N"
                        for p in pats:
                            m = re.findall(p, h_up)
                            if m: nf_res = m[0]; break
                        
                        # Regra de Sinais conforme sua instrução
                        v_deb, v_cre = (-deb, cre) if tipo_robo == "Fornecedor" else (deb, -cre)
                        dados_lista.append({"Data": str(lin[0]), "Historico": hist, "NF_AJUSTADA": nf_res, "Debito": v_deb, "Credito": v_cre})

            df_final = pd.DataFrame(dados_lista)

            if not df_final.empty:
                out = BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    wb = writer.book
                    f_cab = wb.add_format({'bold': 1, 'bg_color': '#9B8ADE', 'font_color': 'white', 'border': 1, 'align': 'center'})
                    f_m = wb.add_format({'num_format': '#,##0.00', 'border': 1})
                    f_c = wb.add_format({'border': 1, 'align': 'center'})
                    
                    # ABA 1: RAZÃO (DADOS)
                    df_final.to_excel(writer, sheet_name='Razao', index=False)
                    ws1 = writer.sheets['Razao']
                    ws1.set_column('A:E', 20)
                    ws1.set_column('B:B', 45) # Coluna Histórico mais larga

                    # ABA 2: CONCILIAÇÃO (RESUMO)
                    # Criamos um resumo único para as notas
                    df_resumo = df_final.groupby("NF_AJUSTADA").agg({"Debito":"sum", "Credito":"sum"}).reset_index()
                    df_resumo["Diferença"] = df_resumo["Debito"] + df_resumo["Credito"]
                    
                    ws2 = wb.add_worksheet('Conciliacao')
                    ws2.set_column('B:F', 20)
                    
                    headers = ["NF_AJUSTADA", "Soma Débito", "Soma Crédito", "Diferença", "Status"]
                    for ci, v in enumerate(headers): ws2.write(1, ci+1, v, f_cab)
                    
                    for ri, r in enumerate(df_resumo.values):
                        row = 2 + ri
                        ws2.write(row, 1, r[0], f_c)
                        ws2.write_number(row, 2, r[1], f_m)
                        ws2.write_number(row, 3, r[2], f_m)
                        ws2.write_number(row, 4, r[3], f_m)
                        status = "OK" if abs(r[3]) < 0.01 else "EM ABERTO"
                        ws2.write(row, 5, status, f_c)

                st.success("✅ Tudo pronto! Agora sem erros.")
                st.download_button("📥 Baixar Relatório SOLUX", out.getvalue(), "solux_oficial.xlsx")
        except Exception as e:
            st.error(f"Erro inesperado: {e}")
