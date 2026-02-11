import streamlit as st
import pandas as pd
import re
from io import BytesIO

# 1. Configuração da Página
st.set_page_config(page_title="SOLUX 2026", page_icon="💡", layout="wide")

# 2. ESTILO SOLUX FINAL
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;800&display=swap');
    .stApp { background-color: #F3F0FF; background-image: url("https://www.transparenttextures.com/patterns/cubes.png"); background-attachment: fixed; }
    header[data-testid="stHeader"], [data-testid="stSidebar"] { background-color: #9B8ADE !important; }
    .titulo { font-family: 'Montserrat', sans-serif; color: #4B0082; font-size: 28px; font-weight: 800; text-align: center; padding: 10px; background-color: rgba(230, 224, 255, 0.9); border-radius: 10px; border: 1px solid #9B8ADE; margin-top: -35px; margin-bottom: 25px; }
    [data-testid="stFileUploaderDropzone"] { background-color: rgba(255, 255, 255, 0.6) !important; border: 2px dashed #9B8ADE !important; }
    [data-testid="stFileUploaderDropzone"] button { background-color: #E6E0FF !important; color: #4B0082 !important; border: 1px solid #9B8ADE !important; font-weight: bold !important; }
    [data-testid="stSidebar"] * { color: #FFFFFF !important; font-weight: 600 !important; }
    .stDownloadButton button { background-color: #9B8ADE !important; color: white !important; border-radius: 8px !important; }
    </style>
    <p class="titulo">💡 SOLUX 2026: Versão Multi-Conciliação Dinâmica 💡</p>
    """, unsafe_allow_html=True)

def to_num(val):
    try:
        if pd.isna(val) or str(val).strip() == '': return 0.0
        s = str(val).replace('.', '').replace(',', '.')
        return float(re.sub(r'[^-0-9.]', '', s))
    except: return 0.0

with st.sidebar:
    st.header("⚙️ Painel de Controle")
    tipo_robo = st.radio("Este projeto é de:", ["Cliente", "Fornecedor"])
    arquivo = st.file_uploader("Suba o arquivo aqui", type=["xlsx", "xls", "csv"])

if arquivo:
    with st.spinner('O robô SOLUX está processando as fórmulas dinâmicas... 🕵️‍♂️✨'):
        try:
            if arquivo.name.endswith('.csv'):
                df_bruto = pd.read_csv(arquivo, header=None, sep=None, engine='python', encoding='latin-1')
            else:
                df_bruto = pd.read_excel(arquivo, header=None)

            nome_emp = "EMPRESA"
            for i in range(min(15, len(df_bruto))):
                if "Empresa:" in str(df_bruto.iloc[i, 0]):
                    nome_emp = str(df_bruto.iloc[i, 2]); break

            banco, f_info = {}, {}
            f_cod, dados = None, []

            for i in range(len(df_bruto)):
                lin = df_bruto.iloc[i]
                if "Conta:" in str(lin[0]):
                    if f_cod and dados: banco[f_cod] = pd.DataFrame(dados)
                    f_cod = str(lin[1]).strip()
                    f_info[f_cod] = f"{f_cod} - {str(lin[5]) if len(lin) > 5 and pd.notna(lin[5]) else str(lin[2])}"
                    dados = []
                elif len(lin) >= 10 and pd.notna(lin[0]) and any(x in str(lin[0]) for x in ['/', '-']):
                    deb, cre = to_num(lin[8]), to_num(lin[9])
                    if deb != 0 or cre != 0:
                        hist = str(lin[2]).strip()
                        if 'TOTAL' in hist.upper(): continue
                        
                        try: data_formatada = pd.to_datetime(lin[0]).strftime('%d/%m/%Y')
                        except: data_formatada = str(lin[0])

                        h_up = hist.upper()
                        pats = [r'SERVIÇO\s?PRESTADO\s?(\d+)', r'NF\s?DE\s?S\s?(\d+)', r'FRETE\s?TOMADO\s?(\d+)', r'CTE\s?(\d+)', r'NFE\s?(\d+)', r'SAÍDA\s?(\d+)', r'NF\s?(\d+)']
                        nf_res = None
                        for p in pats:
                            m = re.findall(p, h_up)
                            if m: nf_res = m[0]; break
                        
                        nf = nf_res if nf_res else "S/ N° NF"
                        v_deb, v_cre = (-deb, cre) if tipo_robo == "Fornecedor" else (deb, -cre)
                        
                        dados.append({"Data": data_formatada, "NF": nf, "Hist": hist, "Deb": v_deb, "Cred": v_cre, "Aviso": (nf == "S/ N° NF")})

            if f_cod and dados: banco[f_cod] = pd.DataFrame(dados)

            if banco:
                out = BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    wb = writer.book
                    # ESTILOS MANTIDOS
                    f_cab = wb.add_format({'bold': 1, 'bg_color': '#E6E0FF', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                    f_emp = wb.add_format({'bold': 1, 'font_size': 14, 'align': 'center', 'bg_color': '#D3D3D3', 'border': 1})
                    f_c = wb.add_format({'align': 'center', 'border': 1})
                    f_m = wb.add_format({'num_format': '#,##0.00', 'border': 1})
                    f_vde = wb.add_format({'num_format': '#,##0.00', 'font_color': 'green', 'bold': 1, 'border': 1})
                    f_vrm = wb.add_format({'num_format': '#,##0.00', 'font_color': 'red', 'bold': 1, 'border': 1})

                    for cod, df in banco.items():
                        ws = wb.add_worksheet(str(cod)[:31])
                        ws.hide_gridlines(2)
                        
                        # Colunas: A(2), B(12), C(12), D(45), E(15), F-G(18), H-I(2.14), J-M(18)
                        ws.set_column('A:A', 2); ws.set_column('B:C', 12); ws.set_column('D:D', 45)
                        ws.set_column('E:E', 15); ws.set_column('F:G', 18); ws.set_column('H:I', 2.14)
                        ws.set_column('J:M', 18); ws.set_column('N:N', 15)

                        ws.merge_range('B2:N2', f"EMPRESA: {nome_emp} ({tipo_robo})", f_emp)
                        ws.merge_range('B4:G4', f_info[cod], f_cab)
                        ws.merge_range('J4:N4', "CONCILIAÇÃO POR NOTA (DINÂMICA)", f_cab)

                        # Tabela Razão com Coluna Nova (NF_AJUSTADA na Coluna E)
                        headers_razao = ["Data", "NF Original", "Histórico", "NF_AJUSTADA", "Débito", "Crédito"]
                        for ci, v in enumerate(headers_razao): ws.write(5, ci+1, v, f_cab)
                        
                        for ri, r in enumerate(df.values):
                            ws.write(6+ri, 1, r[0], f_c) # Data
                            ws.write(6+ri, 2, r[1], f_c) # NF Original
                            ws.write(6+ri, 3, r[2], f_c) # Histórico
                            ws.write(6+ri, 4, r[1], f_c) # NF_AJUSTADA (Começa igual a NF original)
                            ws.write_number(6+ri, 5, r[3], f_m) # Débito
                            ws.write_number(6+ri, 6, r[4], f_m) # Crédito
                        
                        last_row = 6 + len(df)
                        
                        # Tabela Conciliação com FÓRMULAS (Lado Direito)
                        nfs_unicas = df["NF"].unique()
                        headers_conc = ["NF", "Deb (Soma)", "Cred (Soma)", "Diferença", "Status"]
                        for ci, v in enumerate(headers_conc): ws.write(5, ci+9, v, f_cab)
                        
                        for ri, nf in enumerate(nfs_unicas):
                            curr_row = 7 + ri
                            ws.write(curr_row-1, 9, nf, f_c)
                            
                            # Fórmulas SOMASE: Olha para a coluna E (NF_AJUSTADA) e soma F (Deb) e G (Cred)
                            # Se você mudar a NF na coluna E, o resultado aqui muda sozinho!
                            range_ajuste = f"$E$7:$E${last_row}"
                            ws.write_formula(curr_row-1, 10, f'=SUMIF({range_ajuste}, J{curr_row}, $F$7:$F${last_row})', f_m)
                            ws.write_formula(curr_row-1, 11, f'=SUMIF({range_ajuste}, J{curr_row}, $G$7:$G${last_row})', f_m)
                            ws.write_formula(curr_row-1, 12, f'=K{curr_row}+L{curr_row}', f_m)
                            ws.write_formula(curr_row-1, 13, f'=IF(ABS(M{curr_row})<0.01, "OK", "EM ABERTO")', f_c)

                        # Saldos Finais
                        ws.write(last_row + 1, 5, "Saldo Razão:", f_cab)
                        ws.write_formula(last_row + 1, 6, f'=SUM(F7:F{last_row})+SUM(G7:G{last_row})', f_vde)

                st.success(f"✅ Versão Dinâmica Gerada! Coluna E é para seus ajustes.")
                st.download_button("📥 Baixar Relatório SOLUX", out.getvalue(), f"solux_dinamico_{tipo_robo.lower()}.xlsx")
        except Exception as e:
            st.error(f"Erro ao processar: {e}")
