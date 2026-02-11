import streamlit as st
import pandas as pd
import re
from io import BytesIO

# 1. Configuração da Página
st.set_page_config(page_title="SOLUX", page_icon="💡", layout="wide")

# 2. ESTILO (Lavanda Suave e Tabela Elegante)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;800&display=swap');
    .stApp { background-color: #F3F0FF; background-image: url("https://www.transparenttextures.com/patterns/cubes.png"); background-attachment: fixed; }
    header[data-testid="stHeader"], [data-testid="stSidebar"] { background-color: #9B8ADE !important; }
    button[kind="headerNoPadding"], .stApp header svg { display: none !important; }
    .titulo { font-family: 'Montserrat', sans-serif; color: #4B0082; font-size: 28px; font-weight: 800; text-align: center; padding: 8px; background-color: rgba(230, 224, 255, 0.9); border-radius: 10px; border: 1px solid #9B8ADE; margin-top: -35px; margin-bottom: 25px; }
    [data-testid="stSidebar"] * { font-family: 'Montserrat', sans-serif; color: #FFFFFF !important; font-weight: 600 !important; }
    
    /* Estilo para a Tabela Viva */
    .stDataFrame { background-color: white; border-radius: 10px; padding: 10px; box-shadow: 0px 4px 10px rgba(0,0,0,0.05); }
    
    /* Botão de Atualizar */
    .stButton button { background-color: #4B0082 !important; color: white !important; font-weight: bold !important; border-radius: 8px !important; width: 100%; }
    </style>
    <p class="titulo">💡 SOLUX: FINANÇAS INTELIGENTES (EDIÇÃO VIVA) 💡</p>
    """, unsafe_allow_html=True)

def to_num(val):
    try:
        if pd.isna(val) or str(val).strip() == '': return 0.0
        return float(str(val).replace('.', '').replace(',', '.'))
    except: return 0.0

# 4. PAINEL LATERAL
with st.sidebar:
    st.header("⚙️ Painel")
    tipo_robo = st.radio("Este projeto é de:", ["Cliente", "Fornecedor"])
    arquivo = st.file_uploader("Suba o arquivo original aqui", type=["xlsx", "xls", "csv"])

# 5. LÓGICA DO CORAÇÃO (SESSION STATE)
# Isso faz o robô não "esquecer" o que você digitou na tabela
if arquivo:
    if 'df_viva' not in st.session_state:
        with st.spinner('💎 O SOLUX está preparando a sua mesa de trabalho...'):
            try:
                if arquivo.name.endswith('.csv'):
                    df_bruto = pd.read_csv(arquivo, header=None, sep=None, engine='python', encoding='latin-1')
                else:
                    df_bruto = pd.read_excel(arquivo, header=None)
                
                # Processamento inicial (O mesmo que já fazíamos)
                dados_lista = []
                f_cod = "Geral"
                for i in range(len(df_bruto)):
                    lin = df_bruto.iloc[i]
                    if len(lin) > 9 and pd.notna(lin[0]):
                        deb, cre = to_num(lin[8]), to_num(lin[9])
                        if deb != 0 or cre != 0:
                            hist = str(lin[2]).strip()
                            # Regra de busca automática inicial
                            busca_nf = re.findall(r'(?:NF|NFE|NOTA|Nf|nfe)\s?(\d+)', hist)
                            nf = busca_nf[0] if busca_nf else (str(lin[1]).strip() if pd.notna(lin[1]) else "S/N")
                            
                            # Suas regras de sinal (Salvas na memória)
                            if tipo_robo == "Fornecedor": v_deb, v_cre = -deb, cre
                            else: v_deb, v_cre = deb, -cre
                            
                            dados_lista.append({"Data": str(lin[0]), "NF": nf, "Hist": hist, "Deb": v_deb, "Cred": v_cre})
                
                st.session_state.df_viva = pd.DataFrame(dados_lista)
            except Exception as e:
                st.error(f"Erro ao ler arquivo: {e}")

    # EXIBIÇÃO DA TABELA VIVA
    if 'df_viva' in st.session_state:
        st.subheader("📝 Tabela de Ajustes")
        st.info("✍️ Edite o 'Hist' abaixo e clique em atualizar para vincular notas faltantes.")
        
        # A TABELA EDITÁVEL
        df_editado = st.data_editor(
            st.session_state.df_viva,
            use_container_width=True,
            hide_index=True,
            column_config={
                "NF": st.column_config.TextColumn("Nota Fiscal", help="O robô atualizará isso sozinho"),
                "Hist": st.column_config.TextColumn("Histórico (EDITÁVEL)", help="Escreva 'NFE 123' aqui")
            }
        )

        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("🔄 ATUALIZAR CONCILIAÇÃO"):
                # O DETETIVE PASSA DE NOVO NAS LINHAS EDITADAS
                for index, row in df_editado.iterrows():
                    # Procura palavras-chave no que você escreveu
                    achou = re.findall(r'(?:NF|NFE|NOTA|Nf|nfe|Nº|N)\s?(\d+)', str(row['Hist']).upper())
                    if achou:
                        df_editado.at[index, 'NF'] = achou[0]
                
                st.session_state.df_viva = df_editado
                st.success("✨ SOLUX atualizou os vínculos!")
                st.balloons()

        with col2:
            # Gerar Excel com os dados novos da tabela
            out = BytesIO()
            with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                df_editado.to_excel(writer, index=False, sheet_name='Conciliado')
            
            st.download_button("📥 BAIXAR EXCEL AJUSTADO 💎", out.getvalue(), "solux_final.xlsx")

else:
    st.info("👋 Olá! Suba o arquivo e use a Tabela Viva para correções rápidas.")
