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
    .stApp { background-color: #F3F0FF; background-image: url("https://www.transparenttextures.com/patterns/cubes.png"); background-attachment: fixed; }
    header[data-testid="stHeader"], [data-testid="stSidebar"] { background-color: #9B8ADE !important; }
    .titulo { font-family: 'Montserrat', sans-serif; color: #4B0082; font-size: 28px; font-weight: 800; text-align: center; padding: 10px; background-color: rgba(230, 224, 255, 0.9); border-radius: 10px; border: 1px solid #9B8ADE; margin-bottom: 25px; }
    [data-testid="stSidebar"] * { color: #FFFFFF !important; font-weight: 600 !important; }
    .stDownloadButton button { background-color: #9B8ADE !important; color: white !important; border-radius: 8px !important; width: 100%; height: 50px; font-weight: bold !important; }
    </style>
    <p class="titulo">💡 SOLUX: Seu parceiro na conciliação 💡</p>
    """, unsafe_allow_html=True)

# 3. A LUPA DETETIVE (Prioridade para os termos que você definiu)
def extrair_nf_prioritario(historico, documento):
    h = str(historico).upper()
    
    # Lista de busca com as palavras que você pediu
    termos = [
        r'DACTE-D\s?(\d+)', r'DACTE\s?(\d+)', r'CTE\s?(\d+)',
        r'N\.\s?COMPRA\s?(\d+)', r'NF\s?DE\s?S\s?(\d+)', 
        r'SAIDA\s?(\d+)', r'PRESTADO\s?(\d+)', # Adicionado Saída e Prestado
        r'NFE\s?(\d+)', r'NF\s?(\d+)'
    ]
    
    for p in termos:
        achado = re.findall(p, h)
        if achado: return str(int(achado[0])) # Se achar o número do frete ou compra, usa ele na hora!

    # Se não achar nada acima, tenta o número do documento da coluna ao lado
    d = str(documento).strip()
    if d and d.isdigit() and d != '0': return str(int(d))
    
    return "SEM NF"

# 4. Funções de Apoio
def to_num(val):
    if pd.isna(val) or str(val).strip() == '': return 0.0
    try:
        s = str(val).replace('.', '').replace(',', '.').strip()
        return float(s)
    except: return 0.0

# 5. Painel Lateral
with st.sidebar:
    st.header("⚙️ Painel de Controle")
    tipo_projeto = st.radio("Este projeto é de:", ["Cliente", "Fornecedor"])
    st.markdown("---")
    arquivo = st.file_uploader("Suba o arquivo aqui", type=["xlsx", "xls", "csv"])

# 6. Lógica Principal
if arquivo:
    with st.spinner('O robô está procurando as notas... 🕵️‍♂️'):
        try:
            df_bruto = pd.read_excel(arquivo, header=None) if not arquivo.name.endswith('.csv') else pd.read_csv(arquivo, header=None, sep=None, engine='python', encoding='latin-1')
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
                    f_info[f_atual] = f"{f_atual} - {str(lin[5]) if pd.notna(lin[5]) else str(lin[2])}"
                    dados = []
                    continue

                if len(lin) >= 10 and f_atual and re.search(r'\d{2}/\d{2}', c0):
                    hist = str(lin[2]).strip()
                    if "TOTAL" in hist.upper(): continue
                    
                    # Usa a nova lupa que prioriza DACTE/N. COMPRA
                    nf_final = extrair_nf_prioritario(hist, lin[1])
                    
                    deb, cre = to_num(lin[8]), to_num(lin[9])

                    # Regra de Sinais
                    if tipo_projeto == "Fornecedor":
                        v_deb, v_cre = -deb, cre # Fornecedor: Crédito (+) Débito (-)
                    else:
                        v_deb, v_cre = deb, -cre # Cliente: Crédito (-) Débito (+)

                    dados.append({"Data": c0, "NF": nf_final, "Hist": hist, "Deb": v_deb, "Cred": v_cre})

            if f_atual and dados: banco[f_atual] = pd.DataFrame(dados)

            if banco:
                out = BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    for cod, df in banco.items():
                        ws = writer.book.add_worksheet(str(cod)[:31])
                        # Formatos Coloridos
                        f_cab = writer.book.add_format({'bold': 1, 'bg_color': '#9B8ADE', 'font_color': 'white', 'border': 1, 'align': 'center'})
                        f_moeda = writer.book.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
                        f_pos = writer.book.add_format({'num_format': 'R$ #,##0.00', 'border': 1, 'font_color': 'green'})
                        f_neg = writer.book.add_format({'num_format': 'R$ #,##0.00', 'border': 1, 'font_color': 'red'})

                        # Escreve tabelas
                        df.to_excel(writer, sheet_name=str(cod)[:31], startrow=5, startcol=1, index=False)
                        
                        # Resumo Agrupado (Aqui ele vai juntar o DACTE 330 com o pagamento!)
                        resumo = df.groupby("NF").agg({"Deb":"sum","Cred":"sum"}).reset_index()
                        resumo["Diferença"] = resumo["Deb"] + resumo["Cred"]
                        
                        row_res = 5
                        for c_idx, h in enumerate(["NF", "Soma Deb", "Soma Cred", "Diferença"]):
                            ws.write(row_res, c_idx + 8, h, f_cab)
                        
                        for r_idx, r in resumo.iterrows():
                            ws.write(row_res + 1 + r_idx, 8, r['NF'])
                            ws.write(row_res + 1 + r_idx, 9, r['Deb'], f_moeda)
                            ws.write(row_res + 1 + r_idx, 10, r['Cred'], f_moeda)
                            ws.write(row_res + 1 + r_idx, 11, r['Diferença'], f_pos if abs(r['Diferença']) < 0.01 else f_neg)

                st.success("✨ Conciliação ajustada! Agora o robô prioriza o número do Frete/Serviço.")
                st.download_button("📥 Baixar Relatório Corrigido", out.getvalue(), "conciliacao_solux.xlsx")
        except Exception as e:
            st.error(f"Erro: {e}")
