import streamlit as st
import pandas as pd
import re
from io import BytesIO

# 1. Configuração da Página
st.set_page_config(page_title="SOLUX", page_icon="💡", layout="wide")

# 2. ESTILO DA INTERFACE (Tema Visual da SOLUX)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;800&display=swap');
    .stApp { background-color: #F3F0FF; }
    header[data-testid="stHeader"], [data-testid="stSidebar"] { background-color: #9B8ADE !important; }
    .titulo { font-family: 'Montserrat', sans-serif; color: #4B0082; font-size: 28px; font-weight: 800; text-align: center; padding: 10px; background-color: rgba(230, 224, 255, 0.9); border-radius: 10px; border: 1px solid #9B8ADE; margin-top: -35px; margin-bottom: 25px; }
    .stDownloadButton button { background-color: #9B8ADE !important; color: white !important; font-weight: bold; width: 100%; height: 50px; }
    </style>
    <p class="titulo">💡 SOLUX: Seu parceiro na conciliação 💡</p>
    """, unsafe_allow_html=True)

# 3. FUNÇÃO DE LIMPEZA DE NÚMEROS (Ajustada para não errar vírgulas)
def to_num(val):
    try:
        if isinstance(val, (int, float)): return float(val)
        if pd.isna(val) or str(val).strip() == '': return 0.0
        
        s = str(val).strip()
        # Se tem ponto de milhar e vírgula decimal (ex: 1.250,50)
        if ',' in s and '.' in s:
            s = s.replace('.', '').replace(',', '.')
        # Se só tem vírgula (ex: 1250,50)
        elif ',' in s:
            s = s.replace(',', '.')
            
        # Remove R$, espaços e outros símbolos, mantendo apenas números, pontos e o sinal de menos
        limpo = re.sub(r'[^-0-9.]', '', s)
        return float(limpo)
    except: 
        return 0.0

# 4. SIDEBAR E UPLOAD
with st.sidebar:
    st.header("⚙️ Painel de Controle")
    tipo_robo = st.radio("Este projeto é de:", ["Clientes", "Fornecedores"])
    arquivo = st.file_uploader("Suba o arquivo aqui", type=["xlsx", "xls", "csv"])

# 5. PROCESSAMENTO DO ARQUIVO
if arquivo:
    with st.spinner('Conciliando... 🕵️‍♂️✨'):
        try:
            if arquivo.name.endswith('.csv'):
                df_bruto = pd.read_csv(arquivo, header=None, sep=None, engine='python', encoding='latin-1')
            else:
                df_bruto = pd.read_excel(arquivo, header=None)

            nome_emp = "EMPRESA"
            # Tenta achar o nome da empresa nas primeiras linhas
            for i in range(min(15, len(df_bruto))):
                if "Empresa:" in str(df_bruto.iloc[i, 0]):
                    nome_emp = str(df_bruto.iloc[i, 2])
                    break

            banco, f_info = {}, {}
            f_cod, dados = None, []

            for i in range(len(df_bruto)):
                lin = df_bruto.iloc[i]
                
                # Identifica início de uma nova conta/cliente/fornecedor
                if "Conta:" in str(lin[0]):
                    if f_cod and dados: 
                        banco[f_cod] = pd.DataFrame(dados)
                    f_cod = str(lin[1]).strip()
                    # Pega o nome do fornecedor/cliente (geralmente na coluna 5 ou 2)
                    nome_aux = str(lin[5]) if len(lin) > 5 and pd.notna(lin[5]) else str(lin[2])
                    f_info[f_cod] = f"{f_cod} - {nome_aux}"
                    dados = []
                
                # Identifica linhas de lançamento (que tem data com / ou -)
                elif len(lin) >= 10 and pd.notna(lin[0]) and any(x in str(lin[0]) for x in ['/', '-']):
                    deb, cre = to_num(lin[8]), to_num(lin[9])
                    
                    if deb != 0 or cre != 0:
                        hist = str(lin[2]).strip()
                        if 'TOTAL' in hist.upper(): continue
                        
                        try: 
                            data_formatada = pd.to_datetime(lin[0]).strftime('%d/%m/%Y')
                        except: 
                            data_formatada = str(lin[0])

                        # Busca número da Nota Fiscal no histórico
                        pats = [
                            r'SERVIÇO\s?TOMADO\s?(\d+)', r'FRETE\s?TOMADO\s?(\d+)', 
                            r'NF\s?DE\s?S\s?(\d+)', r'CTE\s?(\d+)', r'SAÍDA\s?(\d+)', 
                            r'PRESTADO\s?(\d+)', r'NFE\s?(\d+)', r'NF\s?(\d+)'
                        ]
                        nf_res = None
                        for p in pats:
                            m = re.findall(p, hist.upper())
                            if m: 
                                nf_res = m[0]
                                break
                        
                        nf = nf_res if nf_res else "S/ N° NF"
                        
                        # APLICA A SUA REGRA DE OURO:
                        # Fornecedor: Debito (-), Credito (+)
                        # Cliente: Debito (+), Credito (-)
                        if tipo_robo == "Fornecedores":
                            v_deb, v_cre = -deb, cre
                        else:
                            v_deb, v_cre = deb, -cre
                            
                        dados.append({
                            "Data": data_formatada, 
                            "NF": nf, 
                            "Hist": hist, 
                            "Deb": v_deb, 
                            "Cred": v_cre, 
                            "Aviso": (nf == "S/ N° NF")
                        })

            # Salva o último grupo processado
            if f_cod and dados: 
                banco[f_cod] = pd.DataFrame(dados)

            # 6. EXPORTAÇÃO PARA EXCEL
            if banco:
                out = BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    for cod, df_final in banco.items():
                        # Nome da aba limitado a 31 caracteres (regra do Excel)
                        aba = f_info[cod][:31].replace('/', '-').replace('[', '').replace(']', '')
                        df_final.to_excel(writer, sheet_name=aba, index=False)
                        
                        # Ajuste de largura das colunas
                        worksheet = writer.sheets[aba]
                        for idx, col in enumerate(df_final.columns):
                            worksheet.set_column(idx, idx, 15)

                st.success(f"✅ Conciliação da empresa {nome_emp} concluída!")
                st.download_button(
                    label="📥 Baixar Planilha Conciliada",
                    data=out.getvalue(),
                    file_name=f"Conciliacao_{tipo_robo}_{nome_emp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("⚠️ Nenhum dado de lançamento encontrado no arquivo.")

        except Exception as e:
            st.error(f"Erro ao processar: {e}")

else:
    st.info("👋 Olá! Por favor, suba o arquivo do extrato para começarmos.")
