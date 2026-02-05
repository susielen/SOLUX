import streamlit as st
import pandas as pd
import re
from io import BytesIO

# 1. Configuração da Página
st.set_page_config(page_title="SOLUX", page_icon="💡", layout="wide")

# 2. O ESTILO (Cores Lilás e Roxo Restauradas)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;800&display=swap');
    .stApp { background-color: #F3F0FF; }
    .titulo { font-family: 'Montserrat', sans-serif; color: #4B0082; font-size: 28px; font-weight: 800; text-align: center; padding: 10px; background-color: rgba(230, 224, 255, 0.9); border-radius: 10px; border: 1px solid #9B8ADE; margin-bottom: 25px; }
    [data-testid="stSidebar"] { background-color: #9B8ADE !important; }
    .stDownloadButton button { background-color: #9B8ADE !important; color: white !important; font-weight: bold !important; width: 100%; border-radius: 8px; }
    </style>
    <p class="titulo">💡 SOLUX: Seu parceiro na conciliação 💡</p>
    """, unsafe_allow_html=True)

# 3. Funções de Suporte
def to_num(val):
    if pd.isna(val) or str(val).strip() == '': return 0.0
    try:
        s = str(val).replace('R$', '').replace('.', '').replace(',', '.').strip()
        return float(s)
    except: return 0.0

# 4. A NOVA LUPA (Foca no número que você quer!)
def extrair_nf_exata(historico):
    h = str(historico).upper()
    # Lista prioritária baseada no que você me ensinou
    termos = [
        r'DACTE-D\s?(\d+)', r'DACTE\s?(\d+)', r'CTE\s?(\d+)',
        r'N\.\s?COMPRA\s?(\d+)', r'NF\s?DE\s?S\s?(\d+)', 
        r'SAIDA\s?(\d+)', r'PRESTADO\s?(\d+)',
        r'NFE\s?(\d+)', r'NF\s?(\d+)'
    ]
    for p in termos:
        achado = re.findall(p, h)
        if achado: return str(int(achado[0]))
    
    # Se não achar palavra-chave, busca o último número de 1 a 6 dígitos no texto
    avulso = re.findall(r'\b(\d{1,6})\b', h)
    if avulso: return str(int(avulso[-1]))
    
    return "SEM NF"

# 5. Painel Lateral
with st.sidebar:
    st.header("⚙️ Painel de Controle")
    tipo_proj = st.radio("Este projeto é de:", ["Cliente", "Fornecedor"])
    st.markdown("---")
    arquivo = st.file_uploader("Suba o arquivo aqui", type=["xlsx", "xls", "csv"])

# 6. Processamento
if arquivo:
    try:
        df_bruto = pd.read_excel(arquivo, header=None)
        df_bruto = df_bruto.dropna(how='all').reset_index(drop=True)
        
        banco, f_info = {}, {}
        f_atual, dados = None, []
        empresa = "EMPRESA"

        for i in range(len(df_bruto)):
            lin = df_bruto.iloc[i]
            c0 = str(lin[0])

            if "Empresa:" in c0: empresa = str(lin[2])
            if "Conta:" in c0:
                if f_atual and dados: banco[f_atual] = pd.DataFrame(dados)
                f_atual = str(lin[1]).strip()
                f_info[f_atual] = f"{f_atual} - {str(lin[5])}"
                dados = []
                continue

            if len(lin) >= 10 and f_atual and re.search(r'\d{2}/\d{2}', c0):
                hist = str(lin[2]).strip()
                if "TOTAL" in hist.upper(): continue
                
                # PEGA O NÚMERO APENAS DO HISTÓRICO (Ignora a coluna 1 que estava errada)
                nf_final = extrair_nf_exata(hist)
                
                v_deb = to_num(lin[8])
                v_cre = to_num(lin[9])

                # REGRA DE SINAIS DO USUÁRIO
                if tipo_proj == "Fornecedor":
                    # Crédito (+) e Débito (-)
                    d, c = -v_deb if v_deb != 0 else 0, v_cre if v_cre != 0 else 0
                else:
                    # Cliente: Crédito (-) e Débito (+)
                    d, c = v_deb if v_deb != 0 else 0, -v_cre if v_cre != 0 else 0

                dados.append({"Data": c0, "NF": nf_final, "Hist": hist, "Deb": d, "Cred": c})

        if f_atual and dados: banco[f_atual] = pd.DataFrame(dados)

        if banco:
            out = BytesIO()
            with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                for cod, df in banco.items():
                    ws = writer.book.add_worksheet(str(cod)[:31])
                    f_moeda = writer.book.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
                    f_pos = writer.book.add_format({'num_format': 'R$ #,##0.00', 'border': 1, 'font_color': 'green'})
                    f_neg = writer.book.add_format({'num_format': 'R$ #,##0.00', 'border': 1, 'font_color': 'red'})
                    f_cab = writer.book.add_format({'bold': 1, 'bg_color': '#9B8ADE', 'font_color': 'white', 'border': 1})

                    df.to_excel(writer, sheet_name=str(cod)[:31], startrow=5, startcol=1, index=False)
                    
                    # CONCILIAÇÃO: Agora o DACTE 330 vai aparecer uma vez só!
                    resumo = df.groupby("NF").agg({"Deb": "sum", "Cred": "sum"}).reset_index()
                    resumo["Dif"] = resumo["Deb"] + resumo["Cred"]
                    
                    resumo.to_excel(writer, sheet_name=str(cod)[:31], startrow=5, startcol=8, index=False)
                    # (Espaço para aplicar cores na coluna Dif no Excel final)

            st.success("✨ Agora sim! O robô aprendeu a ler o número do frete corretamente.")
            st.download_button("📥 Baixar Relatório", out.getvalue(), "conciliacao_final.xlsx")
    except Exception as e:
        st.error(f"Erro: {e}")
