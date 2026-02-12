import streamlit as st
import pandas as pd
import re
from io import BytesIO

# 1. Configuração da Página
st.set_page_config(page_title="SOLUX", page_icon="💡", layout="wide")

# 2. ESTILO DA INTERFACE (Ajustando as cores para ficarem iguais)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;800&display=swap');
    
    .stApp { background-color: #F3F0FF; }
    header[data-testid="stHeader"], [data-testid="stSidebar"] { background-color: #9B8ADE !important; }
    
    .titulo { 
        font-family: 'Montserrat', sans-serif; 
        color: #4B0082; 
        font-size: 28px; 
        font-weight: 800; 
        text-align: center; 
        padding: 10px; 
        background-color: #E6E0FF; 
        border-radius: 10px; 
        border: 1px solid #9B8ADE; 
        margin-top: -35px; 
        margin-bottom: 25px; 
    }
    
    .stDownloadButton button { background-color: #9B8ADE !important; color: white !important; font-weight: bold; width: 100%; height: 50px; }

    /* AQUI ESTÁ A COR QUE VOCÊ PEDIU PARA O QUADRADO DO ARQUIVO */
    [data-testid="stFileUploadDropzone"] {
        background-color: #4B0082 !important; /* Mesma cor do fundo do título */
        border: 2px dashed #9B8ADE !important;
        border-radius: 10px;
    }
    
    /* Deixando as letras escuras para conseguir ler no fundo claro */
    [data-testid="stFileUploadDropzone"] p, [data-testid="stFileUploadDropzone"] span {
        color: #4B0082 !important;
        font-weight: bold;
    }
    </style>
    <p class="titulo">💡 SOLUX: Seu parceiro na conciliação 💡</p>
    """, unsafe_allow_html=True)

def to_num(val):
    try:
        if pd.isna(val) or str(val).strip() == '': return 0.0
        s = str(val).replace('.', '').replace(',', '.')
        return float(re.sub(r'[^-0-9.]', '', s))
    except: return 0.0

with st.sidebar:
    st.header("⚙️ Painel de Controle")
    tipo_robo = st.radio("Este projeto é de:", ["Clientes", "Fornecedores"])
    arquivo = st.file_uploader("Suba o arquivo aqui", type=["xlsx", "xls", "csv"])

if arquivo:
    with st.spinner('Consertando os baldes de tinta... ✨'):
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
                        pats = [
                            r'SERVIÇO\s?TOMADO\s?(\d+)', r'FRETE\s?TOMADO\s?(\d+)', 
                            r'NF\s?DE\s?S\s?(\d+)', r'CTE\s?(\d+)', r'SAÍDA\s?(\d+)', 
                            r'PRESTADO\s?(\d+)', r'NFE\s?(\d+)', r'NF\s?(\d+)'
                        ]
                        nf_res = None
                        for p in pats:
                            m = re.findall(p, h_up)
                            if m: nf_res = m[0]; break
                        
                        nf = nf_res if nf_res else "S/ N° NF"
                        v_deb, v_cre = (-deb, cre) if tipo_robo == "Fornecedores" else (deb, -cre)
                        dados.append({"Data": data_formatada, "NF": nf, "Hist": hist, "Deb": v_deb, "Cred": v_cre, "Aviso": (nf == "S/ N° NF")})

            if f_cod and dados: banco[f_cod] = pd.DataFrame(dados)

            if banco:
                out = BytesIO()
                # Aqui estava o erro 'workbook is not defined', agora está certinho
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    workbook = writer.book # Criando o nome do balde de tinta
                    
                    f_cab = workbook.add_format({'bold': 1, 'bg_color': '#F2F2F2', 'align': 'center', 'border': 1})
                    f_emp = workbook.add_format({'bold': 1, 'font_size': 14, 'align': 'center', 'bg_color': '#D3D3D3', 'border': 1})
                    f_c = workbook.add_format({'align': 'center', 'border': 1})
                    f_m = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
                    f_s = workbook.add_format({'border': 1})
                    f_ama_lin = workbook.add_format({'border': 1, 'bg_color': '#FFFF99'})
                    
                    f_saldo_verde = workbook.add_format({'num_format': '#,##0.00', 'font_color': 'green', 'bg_color': '#D3D3D3', 'bold': 1, 'border': 1, 'align': 'center'})
                    f_saldo_vermelho = workbook.add_format({'num_format': '#,##0.00', 'font_color': 'red', 'bg_color': '#D3D3D3', 'bold': 1, 'border': 1, 'align': 'center'})
                    f_label_saldo = workbook.add_format({'bold': 1, 'bg_color': '#D3D3D3', 'align': 'center', 'border': 1})

                    for cod, df_emp in banco.items():
                        ws = workbook.add_worksheet(str(cod)[:31])
                        ws.hide_gridlines(2)
                        ws.set_column('A:A', 2.14) 
                        ws.set_column('B:C', 15); ws.set_column('D:D', 45); ws.set_column('E:F', 18)
                        ws.set_column('G:H', 2.14); ws.set_column('I:M', 18)
                        
                        ws.merge_range('B2:M2', f"EMPRESA: {nome_emp} ({tipo_robo})", f_emp)
                        ws.merge_range('B4:F4', f_info[cod], f_cab)
                        ws.merge_range('I4:M4', "CONCILIAÇÃO POR NOTA", f_cab)

                        for ci, v in enumerate(["Data","NF","Histórico","Débito","Crédito"]):
                            ws.write(5, ci+1, v, f_cab)
                        
                        row_f = 5
                        for ri, r in enumerate(df_emp.values):
                            fmt = f_ama_lin if r[5] else None
                            ws.write(6+ri, 1, r[0], fmt if fmt else f_c)
                            ws.write(6+ri, 2, r[1], fmt if fmt else f_c)
                            ws.write(6+ri, 3, r[2], fmt if fmt else f_s)
                            # Pequeno ajuste para garantir que a linha amarela pinte tudo
                            cel_fmt = workbook.add_format({'num_format': '#,##0.00', 'border': 1, 'bg_color': '#FFFF99' if r[5] else 'white'})
                            ws.write_number(6+ri, 4, r[3], cel_fmt)
                            ws.write_number(6+ri, 5, r[4], cel_fmt)
                            row_f = 6+ri
                        
                        sl = df_emp["Deb"].sum() + df_emp["Cred"].sum()
                        ws.write(row_f + 3, 4, "Saldo Líquido:", f_label_saldo)
                        ws.write_number(row_f + 3, 5, sl, f_saldo_verde if sl >= 0 else f_saldo_vermelho)

                        res = df_emp.groupby("NF").agg({"Deb":"sum", "Cred":"sum"}).reset_index()
                        res["Dif"] = res["Deb"] + res["Cred"]
                        for ci, v in enumerate(["NF","Deb","Cred","Diferença", "Status"]):
                            ws.write(5, ci+8, v, f_cab)
                        
                        row_res = 5
                        for ri, r in enumerate(res.values):
                            ws.write(6+ri, 8, str(r[0]), f_c)
                            ws.write_number(6+ri, 9, r[1], f_m)
                            ws.write_number(6+ri, 10, r[2], f_m)
                            ws.write_number(6+ri, 11, r[3], f_m)
                            ok = abs(r[3]) < 0.01
                            ws.write(6+ri, 12, "OK" if ok else "EM ABERTO", workbook.add_format({'align':'center','bold':1,'border':1,'font_color':'green' if ok else '#CC7A00'}))
                            row_res = 6+ri
                        
                        sf = res["Dif"].sum()
                        ws.write(row_res + 2, 11, "Saldo Final:", f_label_saldo)
                        ws.write_number(row_res + 2, 12, sf, f_saldo_verde if sf >= 0 else f_saldo_vermelho)

                st.success("Solux! Conciliado com sucesso 😁")
                st.download_button("📥 BAIXAR CONICLIAÇÃO SOLUX", out.getvalue(), "solux_conciliacao.xlsx")
        except Exception as e:
            st.error(f"Erro no Processamento: {e}")
