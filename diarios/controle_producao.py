import smartsheet
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import locale
import os

# Tenta configurar para português para pegar os meses por extenso
try:
    locale.setlocale(locale.LC_TIME, "pt_BR.utf8")
except:
    pass # Caso o sistema não suporte, ele segue o padrão

# ================================
# CONFIGURAÇÕES SMARTSHEET
# ================================
TOKEN = '32M5yHYGUMBRyOTkf3GstuBbJ36Q4T9TvefrX'
ID_PLANILHA = '8431517991653252'

def carregar_dados():
    smart = smartsheet.Smartsheet(TOKEN)
    sheet = smart.Sheets.get_sheet(ID_PLANILHA)
    colunas = [col.title for col in sheet.columns]
    dados = []
    for row in sheet.rows:
        dados.append([cell.value for cell in row.cells])
    return pd.DataFrame(dados, columns=colunas)

# 1. Carregar e Filtrar
df = carregar_dados()
data_input = input("Digite a data (ex: 19/03/2026): ").strip()

# Converter para datetime para extrair o dia e mês por extenso
data_dt = pd.to_datetime(data_input, format='%d/%m/%Y')
data_extenso = data_dt.strftime('%d de %B').lower() # ex: 19 de março

df['Data_Str'] = pd.to_datetime(df['Data']).dt.strftime('%d/%m/%Y')
df_filtrado = df[df['Data_Str'] == data_input].copy()

if df_filtrado.empty:
    print("Nenhum dado encontrado.")
    exit()

# Cálculos
df_filtrado['Quantidade Produzida'] = pd.to_numeric(df_filtrado['Quantidade Produzida'], errors='coerce').fillna(0)
producao_setor = df_filtrado.groupby('Setor')['Quantidade Produzida'].sum().reset_index()
total_dia = df_filtrado['Quantidade Produzida'].sum()

# ... (mantenha a parte inicial de conexão ao Smartsheet)

# ================================
# DASHBOARD A3 - ATUALIZADO
# ================================
fig = make_subplots(
    rows=2, cols=1,
    row_heights=[0.35, 0.65],
    vertical_spacing=0.18, # Aumentei o espaço para caber o título das barras
    specs=[[{"type": "indicator"}], [{"type": "bar"}]]
)

# Título Principal: CONTROLE PRODUÇÃO
fig.add_annotation(
    text="<b>CONTROLE PRODUÇÃO</b>",
    xref="paper", yref="paper",
    x=0.5, y=1.18,              # x=0.5 centraliza na horizontal
    xanchor="center",           # Garante que o centro do texto esteja no 0.5
    showarrow=False,
    font=dict(size=90, color="#002D5F"),
    align="center"
)

# Subtítulo: Dia 19 de março
fig.add_annotation(
    text=f"Dia {data_extenso}",
    xref="paper", yref="paper",
    x=0.5, y=1.08,              # Um pouco abaixo do título principal
    xanchor="center",
    showarrow=False,
    font=dict(size=65, color="#555"),
    align="center"
)

# 2. KPI: Sem vírgula no número
fig.add_trace(
    go.Indicator(
        mode="number",
        value=total_dia,
        title={"text": "<b>Quantidade Produzida no Dia</b>", "font": {"size": 50}},
        number={"font": {"size": 220}, "valueformat": "d"}, # 'd' remove as vírgulas e pontos
    ),
    row=1, col=1
)

# 3. Gráfico de Barras com Novo Título
fig.add_trace(
    go.Bar(
        x=producao_setor['Setor'],
        y=producao_setor['Quantidade Produzida'],
        text=producao_setor['Quantidade Produzida'],
        textposition='outside',
        marker_color='#FE4E10',
        textfont=dict(size=45, family="Arial Black")
    ),
    row=2, col=1
)

# Título específico para o gráfico de barras
fig.add_annotation(
    text="<b>Quantidade Produzida por Setor</b>",
    xref="paper", yref="paper",
    x=0.5, y=0.62, # Posicionado acima das barras
    showarrow=False,
    font=dict(size=45, color="#002D5F"),
    xanchor="center"
)

# Configurações de Layout A3
fig.update_layout(
    template="plotly_white",
    width=3508, 
    height=2480,
    margin=dict(t=400, b=150, l=100, r=100), 
    showlegend=False
)

fig.update_xaxes(tickfont=dict(size=42, family="Arial Black"), row=2, col=1)
fig.update_yaxes(visible=False, row=2, col=1)

# Salvar
arquivo = f"Relatorios/Dashboard_Controle.png"
if not os.path.exists("Relatorios"): os.makedirs("Relatorios")
fig.write_image(arquivo)
print(f"\n[OK] Dashboard gerado: {arquivo}")
