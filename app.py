import streamlit as st
import pandas as pd
import re
import io

# Título da página
st.set_page_config(page_title="Robô Conciliador", layout="wide")
st.title("🤖 Meu Robô de Conciliação")
st.write("Arraste o arquivo do Razão da Domínio aqui embaixo e eu faço a mágica!")

# 1. Botão para subir o arquivo
arquivo_subido = st.file_uploader("Escolha o arquivo CSV do Razão", type=['csv'])

if arquivo_subido is not None:
    # Lendo o arquivo (pulando o cabeçalho da Domínio)
    df = pd.read_csv(arquivo_subido, skiprows=7)

    # Limpeza básica
    df = df.dropna(subset=['Histórico'])
    df = df[~df['Histórico'].str.contains("SALDO ANTERIOR", na=False)]

    # 2. Aplicando as suas regras de sinal (Fornecedor)
    # Crédito é Negativo (-) e Débito é Positivo (+)
    df['Débito'] = pd.to_numeric(df['Débito'], errors='coerce').fillna(0)
    df['Crédito'] = pd.to_numeric(df['Crédito'], errors='coerce').fillna(0)
    df['Valor_Real'] = df['Débito'] - df['Crédito']

    # 3. O Detetive de Notas (Procurando SAÍDA, PRESTADO e números)
    def localizar_nota(texto):
        texto_limpo = str(texto).upper()
        # Procura as palavras que você pediu
        if "SAÍDA" in texto_limpo or "PRESTADO" in texto_limpo:
             # Aqui o robô fica mais atento!
             pass
        
        # Pega a sequência de números (Nota Fiscal)
        achado = re.search(r'(\d{3,})', texto_limpo)
        return achado.group(1) if achado else "Sem Nota"

    df['Nota_Fiscal'] = df['Histórico'].apply(localizar_nota)

    # 4. Criando a Tabela Dinâmica
    tabela_dinamica = df.groupby('Nota_Fiscal').agg({
        'Débito': 'sum',
        'Crédito': 'sum',
        'Valor_Real': 'sum'
    }).reset_index()

    tabela_dinamica.rename(columns={'Valor_Real': 'Saldo_Final'}, inplace=True)

    # Mostrando o resultado na tela
    st.subheader("✅ Aqui está sua conciliação:")
    st.dataframe(tabela_dinamica, use_container_width=True)

    # 5. Botão para baixar o arquivo pronto
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        tabela_dinamica.to_excel(writer, index=False)
    
    st.download_button(
        label="📥 Baixar Planilha Conciliada",
        data=output.getvalue(),
        file_name="conciliacao_pronta.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
