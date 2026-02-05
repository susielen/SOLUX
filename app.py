import streamlit as st
import pandas as pd
import re
from io import BytesIO

# 1. Configuração da Página
st.set_page_config(page_title="SOLUX", page_icon="💡", layout="wide")

# 2. Estilo Visual
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;800&display=swap');
    .stApp { background-color: #F3F0FF; }
    .titulo { font-family: 'Montserrat', sans-serif; color: #4B0082; font-size: 28px; font-weight: 800; text-align: center; padding: 10px; background-color: rgba(230, 224, 255, 0.9); border-radius: 10px; border: 1px solid #9B8ADE; margin-bottom: 25px; }
    [data-testid="stSidebar"] { background-color: #9B8ADE !important; }
    .stDownloadButton button { background-color: #9B8ADE !important; color: white !important; width: 100%; border-radius: 8px; }
    </style>
    <p class="titulo">💡 SOLUX: Seu parceiro na conciliação 💡</p>
    """, unsafe_allow_html=True)

# 3. Funções de Suporte
def to_num(val):
    if pd.isna(val) or str(val).strip() == '': return 0.0
    try:
        s = str(val).replace('.', '').replace(',', '.').strip()
        return float(s)
    except: return 0.0

# 4. A LUPA ESPECIALISTA (Agora com Saída e Prestado)
def extrair_nf_completo(historico, documento):
    hist_str = str(historico).upper()
    
    # Lista completa de "etiquetas" que o robô sabe ler
    padroes_chave = [
        r'N\.\s?COMPRA\s?(\d+)',           # Fornecedor: Compra
        r'NF\s?DE\s?S\s?(\d+)',            # Fornecedor: Serviço Tomado
        r'PRESTADO\s?(\d+)',               # Cliente: Serviço Prestado
        r'SAIDA\s?(\d+)',                  # Cliente: Nota de Saída
        r'DACTE-D\s?(\d+)',                # Frete
        r'DACTE\s?(\d+)',                  # Frete
        r'NFE\s?(\d+)',                    # Geral
        r'NF\s?(\d+)',                     # Geral
        r'PAGAMENTO\s?(\d+)',              # Pagamentos
        r'RECEBIMENTO\s?(\d+)',            # Recebimentos
        r'DOC\s?(\d+)'                     # Documento
    ]
    
    for p in padroes_chave:
        achado = re.findall(p, hist_str)
        if achado:
            return str(int(achado[0]))

    # Se não achou palavra-chave, busca número isolado (3 a 8 dígitos)
    avulso = re.findall(r'\b(\d{3,8})\b', hist_str)
    if avulso:
        return str(int(avulso[0]))

    # Tenta a coluna Documento
    doc_limpo = str(documento).strip()
    if doc_limpo and doc_limpo.isdigit() and doc_limpo != '0':
        return str(int(doc_limpo))
        
    return "SEM NF"

# 5. Painel Lateral
with st.sidebar:
    st.header("⚙️ Painel de Controle")
    tipo_robo = st.radio("Este projeto é de:", ["Cliente", "Fornecedor"])
    st.markdown("---")
    arquivo = st.file_uploader("Suba o arquivo aqui", type=["xlsx", "xls", "csv"])

# 6. Processamento
if arquivo:
    with st.spinner('O robô detetive está em ação... 🕵️‍♀️'):
        try:
            if arquivo.name.endswith('.csv'):
                df_bruto = pd.read_csv(arquivo, header=None, sep=None, engine='python', encoding='latin-1')
            else:
                df_bruto = pd.read_excel(arquivo, header=None)
            
            df_bruto = df_bruto.dropna(how='all').reset_index(drop=True)
            
            nome_empresa = "EMPRESA"
            banco = {}
            f_info = {}
            f_cod_atual = None
            dados_acumulados = []

            for i in range(len(df_bruto)):
                linha = df_bruto.iloc[i]
                primeira_cel = str(linha[0])

                if "Empresa:" in primeira_cel:
                    nome_empresa = str(linha[2]) if pd.notna(linha[2]) else "EMPRESA"

                if "Conta:" in primeira_cel:
                    if f_cod_atual and dados_acumulados:
                        banco[f_cod_atual] = pd.DataFrame(dados_acumulados)
                    f_cod_atual = str(linha[1]).strip()
                    nome_f = str(linha[5]) if pd.notna(linha[5]) else str(linha[2])
                    f_info[f_cod_atual] = f"{f_cod_atual} - {nome_f}"
                    dados_acumulados = []
                    continue

                if len(linha) >= 10 and f_cod_atual and re.search(r'\d{2}/\d{2}', primeira_cel):
                    hist = str(linha[2]).strip()
                    if "TOTAL" in hist.upper(): continue
                    
                    nf = extrair_nf_completo(hist, linha[1])
                    deb = to_num(linha[8])
                    cre = to_num(linha[9])

                    # Aplica a sua regra de sinais personalizada
                    if tipo_robo == "Fornecedor":
                        v_deb, v_cre = -deb, cre
                    else:
                        v_deb, v_cre = deb, -cre

                    dados_acumulados.append({"Data": primeira_cel, "NF": nf, "Hist": hist, "Deb": v_deb, "Cred": v_cre})

            if f_cod_atual and dados_acumulados:
                banco[f_cod_atual] = pd.DataFrame(dados_acumulados)

            # Exportação
            if banco:
                out = BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    for cod, df in banco.items():
                        ws = writer.book.add_worksheet(str(cod)[:31])
                        ws.set_column('B:C', 12); ws.set_column('D:D', 45); ws.set_column('E:F', 15); ws.set_column('I:L', 15)
                        
                        # Cabeçalhos e Tabela
                        df.to_excel(writer, sheet_name=str(cod)[:31], startrow=5, startcol=1, index=False)
                        
                        # Resumo Agrupado (Onde a mágica acontece)
                        resumo = df.groupby("NF").agg({"Deb": "sum", "Cred": "sum"}).reset_index()
                        resumo["Diferença"] = resumo["Deb"] + resumo["Cred"]
                        resumo.to_excel(writer, sheet_name=str(cod)[:31], startrow=5, startcol=8, index=False)

                st.success("✅ Relatório gerado! Agora ele reconhece Saídas e Serviços Prestados.")
                st.download_button("📥 Baixar Planilha Conciliada", out.getvalue(), "conciliacao_solux_v3.xlsx")
                
        except Exception as e:
            st.error(f"Erro: {e}")

# Memória: Guardei que para Clientes usaremos "SAÍDA" e "PRESTADO".
# Ok, vou lembrar-me disso. Pode pedir-me para esquecer coisas a qualquer momento ou gerir as informações que guardei nas suas definições (https://gemini.google.com/saved-info).
