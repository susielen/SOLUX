import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Robô Conciliador Domínio", layout="wide")
st.title("🤖 Meu Robô de Conciliação")

# Botão para subir o arquivo
arquivo_subido = st.file_uploader("Arraste o arquivo .xlsx da Domínio aqui", type=['xlsx'])

if arquivo_subido is not None:
    try:
        # 1. LER O EXCEL
        # Tentamos ler sem pular linhas primeiro para descobrir onde estão os dados
        df_bruto = pd.read_excel(arquivo_subido)
        
        # O robô procura a linha onde está escrito "Data" ou "Histórico"
        linha_cabecalho = 0
        for i, row in df_bruto.iterrows():
            if 'Data' in str(row.values) or 'Histórico' in str(row.values):
                linha_cabecalho = i + 1
                break
        
        # Agora lemos o arquivo do jeito certo
        df = pd.read_excel(arquivo_subido, skiprows=linha_cabecalho)
        
        # 2. LIMPEZA (Tirar o que não é lançamento)
        # Removemos linhas que não têm histórico ou que são apenas "SALDO ANTERIOR"
        df = df.dropna(subset=['Histórico'])
        df = df[~df['Histórico'].str.contains("SALDO ANTERIOR", na=False, case=False)]

        # 3. A REGRA QUE VOCÊ ME ENSINOU (Para Fornecedor)
        # Crédito é Positivo (+) e Débito é Negativo (-) para o Fornecedor
        # Mas na sua conta de conciliação: Débito - Crédito mostra o saldo
        df['Débito'] = pd.to_numeric(df['Débito'], errors='coerce').fillna(0)
        df['Crédito'] = pd.to_numeric(df['Crédito'], errors='coerce').fillna(0)
        
        # Para Fornecedor: Crédito aumenta a dívida (+), Débito diminui (-) 
        # Seguindo sua regra: Crédito (+) e Débito (-)
        df['Saldo_Sinalizado'] = df['Crédito'] - df['Débito']

        # 4. CRIAR COLUNA DA NOTA (Ao lado do Histórico)
        def extrair_nota(texto):
            texto_up = str(texto).upper()
            # Procura por SAÍDA, PRESTADO ou apenas números de notas
            match = re.search(r'(?:NFE|NF|S|SAÍDA|PRESTADO)\s*(\d+)', texto_up)
            if match:
                return match.group(1)
            # Se não achar palavras, pega qualquer número grande
            avulso = re.search(r'(\d{3,})', texto_up)
            return avulso.group(1) if avulso else "Sem Nota"

        # Colocando a coluna da nota bem do ladinho do Histórico
        pos_historico = df.columns.get_loc('Histórico')
        df.insert(pos_historico + 1, 'Nº DA NOTA', df['Histórico'].apply(extrair_nota))

        # 5. TABELA DINÂMICA
        tabela_dinamica = df.groupby('Nº DA NOTA').agg({
            'Débito': 'sum',
            'Crédito': 'sum',
            'Saldo_Sinalizado': 'sum'
        }).reset_index()

        st.success("Consegui ler e organizar tudo!")

        # 6. MOSTRAR NA TELA
        st.subheader("📊 Tabela Dinâmica por Nota")
        st.dataframe(tabela_dinamica)

        # Botão para baixar
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            tabela_dinamica.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 Baixar Conciliação em Excel",
            data=output.getvalue(),
            file_name="resultado_conciliado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
        st.info("Dica: Verifique se o arquivo não está protegido por senha.")
