import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="SOLUX 2026", layout="wide")

# Estilo SOLUX
st.markdown("""
    <style>
    .titulo { color: #4B0082; font-size: 24px; font-weight: bold; text-align: center; padding: 10px; background-color: #E6E0FF; border-radius: 10px; }
    .stDownloadButton button { background-color: #9B8ADE !important; color: white !important; font-weight: bold; width: 100%; height: 50px; }
    </style>
    <div class="titulo">💡 SOLUX 2026: PROCESSAMENTO FORÇADO 💡</div>
    """, unsafe_allow_html=True)

def limpar_valor(v):
    try:
        if pd.isna(v) or str(v).strip() == '': return 0.0
        s = str(v).replace('.', '').replace(',', '.')
        return float(re.sub(r'[^-0-9.]', '', s))
    except: return 0.0

tipo = st.sidebar.radio("Projeto:", ["Cliente", "Fornecedor"])
arquivo = st.sidebar.file_uploader("Suba o arquivo", type=["xlsx", "xls", "csv"])

if arquivo:
    try:
        # 1. LER O ARQUIVO (Ignorando erros de colunas vazias)
        if arquivo.name.endswith('.csv'):
            df = pd.read_csv(arquivo, header=None, sep=None, engine='python', encoding='latin-1')
        else:
            df = pd.read_excel(arquivo, header=None)

        # 2. PROCURAR OS DADOS (Onde quer que estejam)
        dados_finais = []
        for i in range(len(df)):
            linha = df.iloc[i].tolist()
            # Procuramos por uma data em qualquer uma das primeiras 3 colunas
            data_encontrada = False
            for col_idx in range(min(4, len(linha))):
                val_celula = str(linha[col_idx])
                if '/' in val_celula and len(val_celula) >= 8:
                    # Achamos a linha de lançamento!
                    # Normalmente: Data(col 1), Hist(col 3), NF(col 4), Deb(col 5), Cred(col 6) no seu exemplo
                    # Mas vamos usar posições relativas ao que vimos no seu CSV
                    try:
                        data = val_celula
                        hist = str(linha[col_idx+2])
                        nf_original = str(linha[col_idx+1])
                        deb = limpar_valor(linha[col_idx+4])
                        cre = limpar_valor(linha[col_idx+5])
                        
                        if deb != 0 or cre != 0:
                            # Regra de Sinais do seu contexto
                            v_deb, v_cre = (-deb, cre) if tipo == "Fornecedor" else (deb, -cre)
                            
                            # Limpar a NF do Histórico (SAÍDA, PRESTADO, NF)
                            pats = [r'SAÍDA\s?(\d+)', r'PRESTADO\s?(\d+)', r'NF\s?(\d+)', r'NFE\s?(\d+)', r'CTE\s?(\d+)']
                            nf_ajustada = nf_original
                            for p in pats:
                                m = re.findall(p, hist.upper())
                                if m: nf_ajustada = m[0]; break
                            
                            dados_finais.append([data, nf_ajustada, hist, v_deb, v_cre])
                            data_encontrada = True
                            break
                    except: continue
            if data_encontrada: continue

        if dados_finais:
            df_razao = pd.DataFrame(dados_finais, columns=["Data", "NF_AJUSTADA", "Histórico", "Débito", "Crédito"])
            
            # 3. CRIAR A CONCILIAÇÃO LADO A LADO
            resumo = df_razao.groupby("NF_AJUSTADA").agg({"Débito":"sum", "Crédito":"sum"}).reset_index()
            resumo["DIFERENÇA"] = resumo["Débito"] + resumo["Crédito"]
            resumo["STATUS"] = resumo["DIFERENÇA"].apply(lambda x: "OK" if abs(x) < 0.01 else "ABERTO")

            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Escreve tudo na mesma aba (Lado a Lado)
                df_razao.to_excel(writer, sheet_name='CONCILIACAO', index=False, startrow=5)
                resumo.to_excel(writer, sheet_name='CONCILIACAO', index=False, startrow=5, startcol=7)
                
                ws = writer.sheets['CONCILIACAO']
                ws.write('A1', f'EMPRESA: PROJETO {tipo.upper()}', workbook.add_format({'bold':True}))
                ws.write('H5', 'CONCILIAÇÃO POR NOTA', workbook.add_format({'bg_color': '#D9EAD3', 'bold': True}))

            st.success("✅ CONSEGUI! Ficheiro processado.")
            st.download_button("📥 BAIXAR AGORA", output.getvalue(), "solux_processado.xlsx")
        else:
            st.warning("Não encontrei as datas de lançamento. Verifique se o arquivo está no formato correto.")

    except Exception as e:
        st.error(f"Erro no Processamento: {e}")
