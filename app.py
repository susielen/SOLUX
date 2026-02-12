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
    <div class="titulo">💡 SOLUX 2026: LAYOUT CONCILIAÇÃO 💡</div>
    """, unsafe_allow_html=True)

def to_num(val):
    try:
        if pd.isna(val) or str(val).strip() == '': return 0.0
        s = str(val).replace('.', '').replace(',', '.')
        return float(re.sub(r'[^-0-9.]', '', s))
    except: return 0.0

tipo = st.sidebar.radio("Projeto:", ["Cliente", "Fornecedor"])
arquivo = st.sidebar.file_uploader("Suba o arquivo Pasta1", type=["xlsx", "xls", "csv"])

if arquivo:
    try:
        # Lendo o arquivo (O seu CSV tem uma vírgula sobrando no começo)
        if arquivo.name.endswith('.csv'):
            df_bruto = pd.read_csv(arquivo, header=None, engine='python', encoding='latin-1')
        else:
            df_bruto = pd.read_excel(arquivo, header=None)

        dados_razao = []
        # Percorre as linhas procurando onde começa a data
        for i in range(len(df_bruto)):
            linha = df_bruto.iloc[i].tolist()
            # No seu arquivo, a data está na coluna 1 (segunda coluna)
            if len(linha) > 6 and pd.notna(linha[1]) and '/' in str(linha[1]):
                data = str(linha[1])
                hist = str(linha[3])
                # Pega Débito e Crédito das colunas 5 e 6
                deb = to_num(linha[5])
                cre = to_num(linha[6])
                
                if deb != 0 or cre != 0:
                    # Busca a NF no histórico
                    pats = [r'NFE\s?(\d+)', r'NF\s?(\d+)', r'SAÍDA\s?(\d+)', r'PRESTADO\s?(\d+)']
                    nf = "S/N"
                    for p in pats:
                        m = re.findall(p, hist.upper())
                        if m: nf = m[0]; break
                    
                    # Ajuste de sinal
                    v_deb, v_cre = (-deb, cre) if tipo == "Fornecedor" else (deb, -cre)
                    dados_razao.append([data, nf, hist, v_deb, v_cre])

        if dados_razao:
            df_final_razao = pd.DataFrame(dados_razao, columns=["Data", "NF", "Histórico", "Débito", "Crédito"])
            
            # Gerando a Dinâmica (Resumo)
            df_dinamica = df_final_razao.groupby("NF").agg({"Crédito":"sum", "Débito":"sum"}).reset_index()
            df_dinamica["DIFERENÇA"] = df_dinamica["Crédito"] + df_dinamica["Débito"]
            df_dinamica["STATUS"] = df_dinamica["DIFERENÇA"].apply(lambda x: "OK" if abs(x) < 0.01 else "EM ABERTO")

            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # 1. Escreve o Razão na mesma folha (Começa na Coluna B)
                df_final_razao.to_excel(writer, sheet_name='CONCILIACAO', index=False, startrow=6, startcol=1)
                
                # 2. Escreve a Dinâmica na mesma folha (Começa na Coluna I)
                df_dinamica.to_excel(writer, sheet_name='CONCILIACAO', index=False, startrow=6, startcol=8)
                
                workbook = writer.book
                ws = writer.sheets['CONCILIACAO']
                
                # Estilos básicos
                f_tit = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#4B0082'})
                f_cab = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1})
                
                # Cabeçalhos manuais para ficar igual ao seu
                ws.write('B2', f'EMPRESA: PROJETO {tipo.upper()}', f_tit)
                ws.write('I6', 'CONCILIAÇÃO POR NOTA', f_cab)
                
                # Largura das colunas
                ws.set_column('B:B', 12); ws.set_column('C:C', 10); ws.set_column('D:D', 40)
                ws.set_column('E:F', 12); ws.set_column('G:G', 4); ws.set_column('I:M', 12)

            st.success("✅ Arquivo Processado! O botão apareceu.")
            st.download_button("📥 BAIXAR AGORA", output.getvalue(), "solux_lado_a_lado.xlsx")
        else:
            st.warning("Não consegui ler os dados. O formato do arquivo pode ter mudado.")

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
