import streamlit as st
import pandas as pd
import re
from io import BytesIO
import time

# 1. Configuração da Página
st.set_page_config(page_title="SOLUX", page_icon="💡", layout="wide")

# 2. Estilo Visual (Cores e Fontes)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;800&display=swap');
    .stApp { background-color: #F3F0FF; background-image: url("https://www.transparenttextures.com/patterns/cubes.png"); background-attachment: fixed; }
    header[data-testid="stHeader"], [data-testid="stSidebar"] { background-color: #9B8ADE !important; }
    .titulo { font-family: 'Montserrat', sans-serif; color: #4B0082; font-size: 28px; font-weight: 800; text-align: center; padding: 10px; background-color: rgba(230, 224, 255, 0.9); border-radius: 10px; border: 1px solid #9B8ADE; margin-bottom: 25px; }
    [data-testid="stSidebar"] * { color: #FFFFFF !important; font-weight: 600 !important; }
    .stDownloadButton button { background-color: #9B8ADE !important; color: white !important; border-radius: 8px !important; width: 100%; }
    </style>
    <p class="titulo">💡 SOLUX: Seu parceiro na conciliação 💡</p>
    """, unsafe_allow_html=True)

# 3. Funções de Suporte
def to_num(val):
    if pd.isna(val) or str(val).strip() == '': return 0.0
    try:
        # Remove pontos de milhar e troca vírgula por ponto
        s = str(val).replace('.', '').replace(',', '.').strip()
        return float(s)
    except: return 0.0

def extrair_nf(historico, documento):
    # Procura padrões comuns de nota fiscal no histórico
    padrao = re.findall(r'(?:NFe|NF|Duplicata|Título|Doc|Fatura|FT|REC)\s?(\d+)', str(historico), re.IGNORECASE)
    if padrao:
        try: return str(int(padrao[0])) # Remove zeros à esquerda (002036 -> 2036)
        except: return str(padrao[0])
    
    # Se não achou no histórico, tenta a coluna de documento
    doc_limpo = str(documento).strip()
    if doc_limpo and doc_limpo.isdigit() and doc_limpo != '0':
        try: return str(int(doc_limpo))
        except: return doc_limpo
        
    return "SEM NF"

# 4. Painel Lateral
with st.sidebar:
    st.header("⚙️ Painel de Controle")
    tipo_robo = st.radio("Este projeto é de:", ["Cliente", "Fornecedor"])
    st.info(f"Regra Ativa: {'Crédito (-) e Débito (+)' if tipo_robo == 'Cliente' else 'Crédito (+) e Débito (-)'}")
    st.markdown("---")
    arquivo = st.file_uploader("Suba o arquivo (Excel ou CSV)", type=["xlsx", "xls", "csv"])

# 5. Processamento Principal
if arquivo:
    with st.spinner('O robô está organizando as caixinhas... ⏳'):
        df_bruto = None
        try:
            if arquivo.name.endswith('.csv'):
                df_bruto = pd.read_csv(arquivo, header=None, sep=None, engine='python', encoding='latin-1')
            elif arquivo.name.endswith('.xls'):
                df_bruto = pd.read_excel(arquivo, header=None, engine='xlrd')
            else:
                df_bruto = pd.read_excel(arquivo, header=None)
            
            df_bruto = df_bruto.dropna(how='all').reset_index(drop=True)
            
        except Exception as e:
            st.error(f"⚠️ Erro ao ler arquivo: {e}")

    if df_bruto is not None:
        try:
            nome_empresa = "EMPRESA"
            banco = {}
            f_info = {}
            f_cod_atual = None
            dados_acumulados = []

            for i in range(len(df_bruto)):
                linha = df_bruto.iloc[i]
                primeira_cel = str(linha[0])

                # Identifica Nome da Empresa
                if "Empresa:" in primeira_cel:
                    nome_empresa = str(linha[2]) if pd.notna(linha[2]) else nome_empresa

                # Identifica Início de um novo Cliente/Fornecedor
                if "Conta:" in primeira_cel:
                    if f_cod_atual and dados_acumulados:
                        banco[f_cod_atual] = pd.DataFrame(dados_acumulados)
                    
                    f_cod_atual = str(linha[1]).strip()
                    nome_entidade = str(linha[5]) if pd.notna(linha[5]) else (str(linha[2]) if pd.notna(linha[2]) else "NOME INDEFINIDO")
                    f_info[f_cod_atual] = f"{f_cod_atual} - {nome_entidade}"
                    dados_acumulados = []
                    continue

                # Processa linhas de valores (verifica se começa com data DD/MM)
                if len(linha) >= 10 and f_cod_atual:
                    if re.search(r'\d{2}/\d{2}', primeira_cel):
                        hist = str(linha[2]).strip()
                        if "TOTAL" in hist.upper(): continue
                        
                        val_deb = to_num(linha[8])
                        val_cre = to_num(linha[9])

                        if val_deb != 0 or val_cre != 0:
                            # Identifica a NF ou coloca "SEM NF"
                            nf_identificada = extrair_nf(hist, linha[1])

                            # APLICAÇÃO DA REGRA DE SINAIS SOLICITADA
                            if tipo_robo == "Fornecedor":
                                # Fornecedor: Crédito (+) e Débito (-)
                                d_final, c_final = -val_deb, val_cre
                            else:
                                # Cliente: Crédito (-) e Débito (+)
                                d_final, c_final = val_deb, -val_cre

                            dados_acumulados.append({
                                "Data": primeira_cel,
                                "NF": nf_identificada,
                                "Hist": hist,
                                "Deb": d_final,
                                "Cred": c_final
                            })

            # Salva o último grupo
            if f_cod_atual and dados_acumulados:
                banco[f_cod_atual] = pd.DataFrame(dados_acumulados)

            # 6. Criação do Excel Final
            if banco:
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    wb = writer.book
                    
                    # Formatos de Célula
                    f_moeda = wb.add_format({'num_format': 'R$ #,##0.00', 'border': 1, 'font_name': 'Arial', 'font_size': 9})
                    f_cab = wb.add_format({'bold': True, 'bg_color': '#9B8ADE', 'font_color': 'white', 'border': 1, 'align': 'center', 'font_name': 'Arial'})
                    f_tit = wb.add_format({'bold': True, 'font_size': 12, 'bg_color': '#F2F2F2', 'align': 'center', 'border': 1, 'font_name': 'Arial'})
                    f_std = wb.add_format({'border': 1, 'font_name': 'Arial', 'font_size': 9})
                    f_neg = wb.add_format({'num_format': 'R$ #,##0.00', 'border': 1, 'font_color': 'red', 'bold': True})
                    f_pos = wb.add_format({'num_format': 'R$ #,##0.00', 'border': 1, 'font_color': 'green', 'bold': True})

                    for cod, df in banco.items():
                        ws = wb.add_worksheet(str(cod)[:31])
                        ws.hide_gridlines(2)
                        
                        # Largura das Colunas
                        ws.set_column('B:C', 12); ws.set_column('D:D', 40); ws.set_column('E:F', 15); ws.set_column('I:L', 15)
                        
                        # Cabeçalho da Planilha
                        ws.merge_range('B2:F2', f"EMPRESA: {nome_empresa}", f_tit)
                        ws.merge_range('B3:F3', f_info[cod], f_tit)
                        
                        # Tabela de Movimentação (Lado Esquerdo)
                        headers = ["Data", "NF", "Histórico", "Débito", "Crédito"]
                        for c_idx, h in enumerate(headers):
                            ws.write(5, c_idx + 1, h, f_cab)
                        
                        for r_idx, row in df.iterrows():
                            ws.write(6 + r_idx, 1, row['Data'], f_std)
                            ws.write(6 + r_idx, 2, row['NF'], f_std)
                            ws.write(6 + r_idx, 3, row['Hist'], f_std)
                            ws.write(6 + r_idx, 4, row['Deb'], f_moeda)
                            ws.write(6 + r_idx, 5, row['Cred'], f_moeda)

                        # TABELA DE CONCILIAÇÃO (Lado Direito - Agrupado)
                        ws.merge_range('I5:L5', "CONCILIAÇÃO (RESUMO POR NF)", f_cab)
                        resumo = df.groupby("NF", sort=False).agg({"Deb": "sum", "Cred": "sum"}).reset_index()
                        resumo["Saldo"] = resumo["Deb"] + resumo["Cred"]
                        
                        headers_res = ["NF", "Total Déb", "Total Cred", "Diferença"]
                        for c_idx, h in enumerate(headers_res):
                            ws.write(6, c_idx + 8, h, f_cab)

                        for r_idx, row in resumo.iterrows():
                            ws.write(7 + r_idx, 8, row['NF'], f_std)
                            ws.write(7 + r_idx, 9, row['Deb'], f_moeda)
                            ws.write(7 + r_idx, 10, row['Cred'], f_moeda)
                            
                            # Cor condicional: Vermelho se não zerou, Verde se zerou
                            fmt_diff = f_moeda
                            if abs(row['Saldo']) > 0.01: fmt_diff = f_neg
                            elif abs(row['Saldo']) <= 0.01: fmt_diff = f_pos
                            
                            ws.write(7 + r_idx, 11, row['Saldo'], fmt_diff)

                st.success("✨ Relatório lapidado com sucesso!")
                st.download_button("📥 Baixar Planilha Conciliada", output.getvalue(), "conciliacao_solux_final.xlsx")
            else:
                st.warning("🔎 O robô leu o arquivo, mas não encontrou dados de 'Conta:' ou valores válidos.")

        except Exception as e:
            st.error(f"🔥 Erro no processamento dos dados: {e}")

# Instrução visual caso não tenha arquivo
if not arquivo:
    st.info("💡 Por favor, carregue o arquivo no painel lateral para começar.")
