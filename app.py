import streamlit as st
import pandas as pd
import re
from io import BytesIO

# Configuração da Página
st.set_page_config(page_title="SOLUX 2026", layout="wide")

st.markdown("""
    <style>
    .titulo { color: #4B0082; font-size: 26px; font-weight: bold; text-align: center; padding: 10px; background-color: #E6E0FF; border-radius: 10px; }
    .stDownloadButton button { background-color: #9B8ADE !important; color: white !important; font-weight: bold; width: 100%; height: 60px; font-size: 20px; }
    </style>
    <div class="titulo">💡 SOLUX 2026: LEITOR DE EXCEL (XLSX) 💡</div>
    """, unsafe_allow_html=True)

def to_num(val):
    try:
        if pd.isna(val) or str(val).strip() == '': return 0.0
        # Trata números formatados como texto no Excel
        s = str(val).replace('.', '').replace(',', '.')
        return float(re.sub(r'[^-0-9.]', '', s))
    except: return 0.0

tipo = st.sidebar.radio("Projeto:", ["Cliente", "Fornecedor"])
arquivo = st.sidebar.file_uploader("Suba o seu ficheiro XLSX aqui", type=["xlsx"])

if arquivo:
    try:
        # 1. LER O EXCEL (XLSX) USANDO O MOTOR CORRETO
        # O engine='openpyxl' é o segredo para ele não travar
        df_bruto = pd.read_excel(arquivo, header=None, engine='openpyxl')

        dados_razao = []
        
        # Percorre as linhas do Excel
        for i in range(len(df_bruto)):
            linha = df_bruto.iloc[i].tolist()
            
            # Procura a linha que tem a data (geralmente na coluna B, que é o índice 1)
            # Verificamos se há uma data ou algo que pareça data nas primeiras colunas
            if len(linha) > 6:
                celula_data = linha[1]
                # Se for um objeto de data do Excel ou texto com '/'
                if pd.api.types.is_datetime64_any_dtype(type(celula_data)) or '/' in str(celula_data):
                    try:
                        data = celula_data.strftime('%d/%m/%Y') if hasattr(celula_data, 'strftime') else str(celula_data)
                        hist = str(linha[3])
                        deb = to_num(linha[5])
                        cre = to_num(linha[6])
                        
                        if deb != 0 or cre != 0:
                            # Busca NF no histórico (SAÍDA, PRESTADO, NF, NFE)
                            pats = [r'SAÍDA\s?(\d+)', r'PRESTADO\s?(\d+)', r'NF\s?(\d+)', r'NFE\s?(\d+)']
                            nf = "S/N"
                            for p in pats:
                                m = re.findall(p, hist.upper())
                                if m: nf = m[0]; break
                            
                            # Regra de Sinais (Contexto: Fornecedor C+/D- | Cliente C-/D+)
                            v_deb, v_cre = (-deb, cre) if tipo == "Fornecedor" else (deb, -cre)
                            
                            dados_razao.append([data, nf, hist, v_deb, v_cre])
                    except:
                        continue

        if dados_razao:
            df_final_razao = pd.DataFrame(dados_razao, columns=["Data", "NF", "Histórico", "Débito", "Crédito"])
            
            # Criar o resumo (Tabela Dinâmica Manual)
            df_resumo = df_final_razao.groupby("NF").agg({"Débito":"sum", "Crédito":"sum"}).reset_index()
            df_resumo["DIFERENÇA"] = df_resumo["Débito"] + df_resumo["Crédito"]
            df_resumo["STATUS"] = df_resumo["DIFERENÇA"].apply(lambda x: "OK" if abs(x) < 0.01 else "EM ABERTO")

            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Escreve o Razão (Lado Esquerdo)
                df_final_razao.to_excel(writer, sheet_name='CONCILIACAO', index=False, startrow=5, startcol=1)
                
                # Escreve a Dinâmica (Lado Direito)
                df_resumo.to_excel(writer, sheet_name='CONCILIACAO', index=False, startrow=5, startcol=8)
                
                workbook = writer.book
                ws = writer.sheets['CONCILIACAO']
                
                # Títulos e Formatação
                f_tit = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#4B0082'})
                ws.write('B2', f'EMPRESA: PROJETO {tipo.upper()}', f_tit)
                ws.write('I5', 'CONCILIAÇÃO POR NOTA', workbook.add_format({'bold': True, 'bg_color': '#D9EAD3'}))
                
                # Ajuste de colunas
                ws.set_column('B:B', 12); ws.set_column('C:C', 10); ws.set_column('D:D', 40)
                ws.set_column('I:M', 15)

            st.success("✅ Excel XLSX lido com sucesso! O botão apareceu.")
            st.download_button("📥 BAIXAR AGORA", output.getvalue(), "solux_conciliado.xlsx")
        else:
            st.warning("Não encontrei dados de lançamentos no seu Excel. Verifique se as datas estão na coluna B.")

    except Exception as e:
        st.error(f"Erro ao ler o ficheiro XLSX: {e}")
