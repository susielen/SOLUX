import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Robô Conciliador Domínio", layout="wide")
st.title("🤖 Robô de Conciliação (Excel)")

# Botão para subir o arquivo .xlsx
arquivo_subido = st.file_uploader("Arraste seu arquivo do Razão (Excel) aqui", type=['xlsx'])

if arquivo_subido is not None:
    # 1. LER O EXCEL (O robô pula as 7 primeiras linhas de cabeçalho da Domínio)
    df = pd.read_excel(arquivo_subido, skiprows=7)

    # 2. LIMPEZA: Remove linhas vazias e o "Saldo Anterior"
    df = df.dropna(subset=['Histórico'])
    df = df[~df['Histórico'].str.contains("SALDO ANTERIOR", na=False, case=False)]

    # 3. REGRAS DE OURO (Personalizadas para você)
    # Para Fornecedor: Débito é positivo (+) e Crédito é negativo (-)
    df['Débito'] = pd.to_numeric(df['Débito'], errors='coerce').fillna(0)
    df['Crédito'] = pd.to_numeric(df['Crédito'], errors='coerce').fillna(0)
    df['Saldo_Calculado'] = df['Débito'] - df['Crédito']

    # 4. DETETIVE DE NOTAS: Procurando números e palavras-chave
    def extrair_nota(texto):
        texto_up = str(texto).upper()
        
        # O robô prioriza as palavras que você ensinou
        # Mas o foco principal é pegar o número da nota fiscal
        # Busca padrões como "NFe 80009", "NF de S 5", etc.
        busca = re.search(r'(?:NFE|NF|S|SAÍDA|PRESTADO)\s*(\d+)', texto_up)
        if busca:
            return busca.group(1)
        
        # Se não achar as palavras, ele pega qualquer número com 3 ou mais dígitos
        n_avulso = re.search(r'(\d{3,})', texto_up)
        return n_avulso.group(1) if n_avulso else "Verificar"

    df.insert(df.columns.get_loc('Histórico') + 1, 'NÚMERO DA NOTA', df['Histórico'].apply(extrair_nota))

    # 5. TABELA DINÂMICA (Agrupando por Nota)
    tabela_dinamica = df.groupby('NÚMERO DA NOTA').agg({
        'Débito': 'sum',
        'Crédito': 'sum',
        'Saldo_Calculado': 'sum'
    }).reset_index()

    # 6. EXIBIÇÃO
    st.success("Planilha organizada com sucesso!")
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Planilha Detalhada")
        st.dataframe(df[['Data', 'Histórico', 'NÚMERO DA NOTA', 'Débito', 'Crédito', 'Saldo_Calculado']])

    with col2:
        st.subheader("Tabela Dinâmica (Conciliação)")
        st.dataframe(tabela_dinamica)

    # Botão de Download
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        tabela_dinamica.to_excel(writer, index=False, sheet_name='Conciliacao')
    
    st.download_button(
        label="📥 Baixar Tabela Dinâmica Pronta",
        data=output.getvalue(),
        file_name="conciliacao_dominio.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
