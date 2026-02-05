import streamlit as st
import pandas as pd
import re
from io import BytesIO

# 1. Configuração da Página
st.set_page_config(page_title="SOLUX", page_icon="💡", layout="wide")

# 2. O ESTILO (Cores Originais Restauradas)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;800&display=swap');
    .stApp { background-color: #F3F0FF; background-image: url("https://www.transparenttextures.com/patterns/cubes.png"); background-attachment: fixed; }
    header[data-testid="stHeader"], [data-testid="stSidebar"] { background-color: #9B8ADE !important; }
    .titulo { font-family: 'Montserrat', sans-serif; color: #4B0082; font-size: 28px; font-weight: 800; text-align: center; padding: 10px; background-color: rgba(230, 224, 255, 0.9); border-radius: 10px; border: 1px solid #9B8ADE; margin-bottom: 25px; }
    [data-testid="stSidebar"] * { color: #FFFFFF !important; font-weight: 600 !important; }
    .stDownloadButton button { background-color: #9B8ADE !important; color: white !important; border-radius: 8px !important; border: none !important; padding: 10px 20px !important; width: 100%; }
    [data-testid="stFileUploaderDropzone"] { background-color: rgba(255, 255, 255, 0.4) !important; border: 2px dashed #9B8ADE !important; border-radius: 12px !important; }
    </style>
    <p class="titulo">💡 SOLUX: Seu parceiro na conciliação 💡</p>
    """, unsafe_allow_html=True)

# 3. Funções de Suporte (A Máquina de Números)
def to_num(val):
    if pd.isna(val) or str(val).strip() == '': return 0.0
    try:
        s = str(val).replace('.', '').replace(',', '.').strip()
        return float(s)
    except: return 0.0

def extrair_nf(historico, documento):
    hist_str = str(historico).upper()
    # Lista de busca incluindo suas novas palavras: SAÍDA e PRESTADO
    padroes = [
        r'N\.\s?COMPRA\s?(\d+)', r'NF\s?DE\s?S\s?(\d+)', r'DACTE-D\s?(\d+)',
        r'DACTE\s?(\d+)', r'PRESTADO\s?(\d+)', r'SAIDA\s?(\d+)',
        r'NFE\s?(\d+)', r'NF\s?(\d+)', r'PAGAMENTO\s?(\d+)', r'DOC\s?(\d+)'
    ]
    for p in padroes:
        achado = re.findall(p, hist_str)
        if achado: return str(int(achado[0]))
    
    # Se não achou, tenta qualquer número de 3 a 8 dígitos
    avulso = re.findall(r'\b(\d{3,8})\b', hist_str)
    if avulso: return str(int(avulso[0]))
    
    # Tenta a coluna de documento
    doc_limpo = str(documento).strip()
    if doc_limpo and doc_limpo.isdigit(): return str(int(doc_limpo))
    
    return "SEM NF"

# 4. Painel Lateral
with st.sidebar:
    st.header("⚙️ Painel de Controle")
    tipo_robo = st.radio("Este projeto é de:", ["Cliente", "Fornecedor"])
    st.markdown("---")
    arquivo = st.file_uploader("Suba o arquivo aqui", type=["xlsx", "xls", "csv"])

# 5. Lógica de Conciliação
if arquivo:
    with st.spinner('O robô está trabalhando... 🕵️‍♂️'):
        try:
            if arquivo.name.endswith('.csv'):
                df_bruto = pd.read_csv(arquivo, header=None, sep=None, engine='python', encoding='latin-1')
            else:
                df_bruto = pd.read_excel(arquivo, header=None)
            
            df_bruto = df_bruto.dropna(how='all').reset_index(drop=True)
            
            banco, f_info = {}, {}
            f_cod_atual, dados_acumulados = None, []
            nome_empresa = "EMPRESA"

            for i in range(len(df_bruto)):
                lin = df_bruto.iloc[i]
                primeira_cel = str(lin[0])

                if "Empresa:" in primeira_cel:
                    nome_empresa = str(lin[2]) if pd.notna(lin[2]) else "EMPRESA"

                if "Conta:" in primeira_cel:
                    if f_cod_atual and dados_acumulados:
                        banco[f_cod_atual] = pd.DataFrame(dados_acumulados)
                    f_cod_atual = str(lin[1]).strip()
                    nome_f = str(lin[5]) if pd.notna(lin[5]) else str(lin[2])
                    f_info[f_cod_atual] = f"{f_cod_atual} - {nome_f}"
                    dados_acumulados = []
                    continue

                if len(lin) >= 10 and f_cod_atual and re.search(r'\d{2}/\d{2}', primeira_cel):
                    hist = str(lin[2]).strip()
                    if "TOTAL" in hist.upper(): continue
                    
                    nf = extrair_nf(hist, lin[1])
                    deb, cre = to_num(lin[8]), to_num(lin[9])

                    # REGRA DE SINAIS (Conforme nossa conversa)
                    if tipo_robo == "Fornecedor":
                        # Crédito (+) e Débito (-)
                        v_deb, v_cre = -deb, cre
                    else:
                        # Cliente: Crédito (-) e Débito (+)
                        v_deb, v_cre = deb, -cre

                    dados_acumulados.append({"Data": primeira_cel, "NF": nf, "Hist": hist, "Deb": v_deb, "Cred": v_cre})

            if f_cod_atual and dados_acumulados:
                banco[f_cod_atual] = pd.DataFrame(dados_acumulados)

            if banco:
                out = BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    wb = writer.book
                    # Formatos
                    f_moeda = wb.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
                    f_cab = wb.add_format({'bold': 1, 'bg_color': '#9B8ADE', 'font_color': 'white', 'border': 1, 'align': 'center'})
                    f_tit = wb.add_format({'bold': 1, 'bg_color': '#F2F2F2', 'border': 1, 'align': 'center'})
                    f_pos = wb.add_format({'num_format': 'R$ #,##0.00', 'border': 1, 'font_color': 'green'})
                    f_neg = wb.add_format({'num_format': 'R$ #,##0.00', 'border': 1, 'font_color': 'red'})

                    for cod, df in banco.items():
                        ws = wb.add_worksheet(str(cod)[:31])
                        ws.set_column('B:C', 12); ws.set_column('D:D', 45); ws.set_column('E:F', 15); ws.set_column('I:L', 15)
                        
                        ws.merge_range('B2:F2', f"EMPRESA: {nome_empresa}", f_tit)
                        ws.merge_range('B3:F3', f_info[cod], f_tit)
                        
                        # Tabela Principal
                        for c_idx, h in enumerate(["Data","NF","Histórico","Débito","Crédito"]):
                            ws.write(5, c_idx+1, h, f_cab)
                        
                        for r_idx, row in df.iterrows():
                            ws.write(6+r_idx, 1, row['Data']); ws.write(6+r_idx, 2, row['NF'])
                            ws.write(6+r_idx, 3, row['Hist']); ws.write(6+r_idx, 4, row['Deb'], f_moeda)
                            ws.write(6+r_idx, 5, row['Cred'], f_moeda)

                        # Conciliação Agrupada
                        ws.merge_range('I5:L5', "CONCILIAÇÃO POR NOTA", f_cab)
                        resumo = df.groupby("NF").agg({"Deb":"sum","Cred":"sum"}).reset_index()
                        resumo["Dif"] = resumo["Deb"] + resumo["Cred"]
                        
                        for c_idx, h in enumerate(["NF","Soma Déb","Soma Cred","Diferença"]):
                            ws.write(6, c_idx+8, h, f_cab)
                        
                        for r_idx, row in resumo.iterrows():
                            ws.write(7+r_idx, 8, row['NF'])
                            ws.write(7+r_idx, 9, row['Deb'], f_moeda)
                            ws.write(7+r_idx, 10, row['Cred'], f_moeda)
                            ws.write(7+r_idx, 11, row['Dif'], f_pos if abs(row['Dif']) < 0.01 else f_neg)

                st.success("✅ Relatório restaurado e atualizado!")
                st.download_button("📥 Baixar Planilha Conciliada", out.getvalue(), "conciliacao_solux.xlsx")
        except Exception as e:
            st.error(f"⚠️ Houve um probleminha: {e}")
