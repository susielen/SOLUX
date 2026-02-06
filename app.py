import pandas as pd
import re

# 1. Abrir a bagunça (Planilha da Domínio)
df = pd.read_excel('razao_dominio.xlsx')

# 2. Tirar as colunas vazias e desmesclar (o pandas faz isso lendo linha a linha)
df = df.dropna(how='all', axis=1)

# 3. Criar a coluna "n° nf" do lado do Histórico
def extrair_nf(texto):
    texto = str(texto)
    # Procura por números (ex: NF 123 ou apenas o número)
    busca = re.findall(r'\b\d+\b', texto)
    return busca[0] if busca else "sem nf"

# Localiza a coluna de Histórico e cria a nova coluna do lado
idx_hist = df.columns.get_loc('Histórico')
df.insert(idx_hist + 1, 'n° nf', df['Histórico'].apply(extrair_nf))

# 4. Ajustar Crédito e Débito conforme sua regra:
# Para Cliente: Crédito (-) e Débito (+)
# (Aqui o robô segue sua orientação de sinal)
df['Diferença'] = df['Débito'] - df['Crédito']

# 5. Criar a Tabela Dinâmica (Resumo)
tabela_dinamica = df.pivot_table(
    index='n° nf', 
    values=['Débito', 'Crédito', 'Diferença'], 
    aggfunc='sum'
)

# 6. Salvar tudo bonitinho
with pd.ExcelWriter('razao_limpo.xlsx', engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='Dados Limpos')
    # Pula as colunas após o Saldo para colocar a dinâmica
    tabela_dinamica.to_excel(writer, sheet_name='Dados Limpos', startcol=len(df.columns) + 2)

print("Robô: Tarefa concluída! Planilha organizada.")
