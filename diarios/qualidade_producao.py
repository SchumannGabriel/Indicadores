import smartsheet
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os

# ======================================================
# 1. CONFIGURAÇÕES E CORES
# ======================================================
TOKEN = '32M5yHYGUMBRyOTkf3GstuBbJ36Q4T9TvefrX'
ID_PLANILHA = '3432321207193476'

COL_DATA = 'Data do relato'
COL_TRAT = 'Tratativa'
COL_CAUSA = 'Causa'
COL_QTD = 'Quantidade pç NC'

AZUL_PCP = "#002D5F"
LARANJA_PCP = "#FE4E10"
CINZA_REGISTRO = "#555555"

CORES_MAP = {'RETRABALHO': AZUL_PCP, 'SUCATA': LARANJA_PCP, 'REGISTRO': CINZA_REGISTRO}

# ======================================================
# 2. CARREGAMENTO E TRATAMENTO
# ======================================================
def carregar_dados():
    smart = smartsheet.Smartsheet(TOKEN)
    sheet = smart.Sheets.get_sheet(ID_PLANILHA)
    colunas = [col.title for col in sheet.columns]
    dados = [[cell.value for cell in row.cells] for row in sheet.rows]
    df = pd.DataFrame(dados, columns=colunas)
    df.columns = df.columns.str.strip()
    df[COL_DATA] = pd.to_datetime(df[COL_DATA], errors='coerce')
    df[COL_TRAT] = df[COL_TRAT].astype(str).str.upper().str.strip()
    df[COL_QTD] = pd.to_numeric(df[COL_QTD], errors='coerce').fillna(0)
    return df

df = carregar_dados()
data_input = input("Data para análise (dd/mm/aaaa): ").strip()
data_dt = pd.to_datetime(data_input, format='%d/%m/%Y')

df_dia = df[df[COL_DATA].dt.strftime('%d/%m/%Y') == data_input].copy()
df_mes = df[(df[COL_DATA].dt.month == data_dt.month) & (df[COL_DATA].dt.year == data_dt.year) & (df[COL_DATA] <= data_dt)].copy()
df_ano = df[(df[COL_DATA].dt.year == data_dt.year) & (df[COL_DATA] <= data_dt)].copy()

# ======================================================
# 3. PREPARAÇÃO DOS DADOS
# ======================================================
vals_kpi = [len(df_dia), df_dia[COL_QTD].sum(), 
            df_dia[df_dia[COL_TRAT]=='RETRABALHO'][COL_QTD].sum(),
            df_dia[df_dia[COL_TRAT]=='SUCATA'][COL_QTD].sum(),
            df_dia[df_dia[COL_TRAT]=='REGISTRO'][COL_QTD].sum()]
labels_kpi = ["Total Não Conformidades", "Total Peças Não Conforme", "Peças Retrabalhadas", "Peças Sucateadas", "Peças Registradas"]

df_causas = df_dia[df_dia[COL_TRAT] != 'REGISTRO']
causas_count = df_causas.groupby(COL_CAUSA)[COL_QTD].sum().sort_values(ascending=True).tail(5)
causas_curtas = [c[:22] + '...' if len(c) > 22 else c for c in causas_count.index]

diario_grouped = df_mes[df_mes[COL_TRAT].isin(['RETRABALHO', 'SUCATA'])].groupby(
    [df_mes[COL_DATA].dt.day, COL_TRAT])[COL_QTD].sum().unstack(fill_value=0)

mensal_grouped = df_ano.groupby([df_ano[COL_DATA].dt.month, COL_TRAT])[COL_QTD].sum().unstack(fill_value=0)
meses_lista = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
mensal_grouped.index = [meses_lista[int(m)-1] for m in mensal_grouped.index if int(m) <= 12]

# ======================================================
# 4. MONTAGEM DO DASHBOARD
# ======================================================
fig = make_subplots(
    rows=4, cols=2,
    row_heights=[0.08, 0.28, 0.28, 0.36],
    specs=[[{"type": "indicator", "colspan": 2}, None],
           [{"type": "domain"}, {"type": "xy"}],
           [{"type": "xy"}, {"type": "xy"}],
           [{"colspan": 2, "type": "xy"}, None]],
    vertical_spacing=0.07,
    horizontal_spacing=0.15,
    subplot_titles=("", "<b>Representatividade Tratativas</b>", "<b>Top 5 Causas do Dia</b>",
                    "<b>Retrabalho Diário</b>", "<b>Sucata Diária</b>",
                    "<b>Evolução Mensal Acumulada</b>")
)

# 1. KPIs (Colados no topo)
for i in range(5):
    fig.add_trace(go.Indicator(
        mode="number", value=vals_kpi[i],
        title={"text": labels_kpi[i], "font": {"size": 26, "color": "#333"}},
        number={"font": {"size": 95, "color": LARANJA_PCP}},
        domain={'x': [0.2*i, 0.2*i+0.18], 'y': [0.94, 1]}
    ))

# 2. Rosca
counts = df_dia[COL_TRAT].value_counts()
fig.add_trace(go.Pie(
    labels=counts.index, values=counts.values, hole=0.55,
    marker=dict(colors=[CORES_MAP.get(x) for x in counts.index]),
    textinfo='percent', textposition='outside',
    textfont=dict(size=28, family="Arial Black"), showlegend=False
), row=2, col=1)

# 3. Top Causas
fig.add_trace(go.Bar(
    x=causas_count.values, y=causas_curtas, orientation='h',
    marker_color=AZUL_PCP, text=causas_count.values,
    textposition='outside', textfont=dict(size=28, family="Arial Black"),
    showlegend=False
), row=2, col=2)

# 4. Diário: RETRABALHO
if 'RETRABALHO' in diario_grouped.columns:
    fig.add_trace(go.Bar(
        x=diario_grouped.index, y=diario_grouped['RETRABALHO'],
        marker_color=AZUL_PCP, text=diario_grouped['RETRABALHO'],
        textposition='outside', textfont=dict(size=20, family="Arial Black"),
        name="Retrabalho", showlegend=False
    ), row=3, col=1)

# 5. Diário: SUCATA
if 'SUCATA' in diario_grouped.columns:
    fig.add_trace(go.Bar(
        x=diario_grouped.index, y=diario_grouped['SUCATA'],
        marker_color=LARANJA_PCP, text=diario_grouped['SUCATA'],
        textposition='outside', textfont=dict(size=20, family="Arial Black"),
        name="Sucata", showlegend=False
    ), row=3, col=2)

# 6. Mensal (Legenda Única)
for t in ['RETRABALHO', 'SUCATA', 'REGISTRO']:
    if t in mensal_grouped.columns:
        fig.add_trace(go.Bar(
            name=t.capitalize(), x=mensal_grouped.index, y=mensal_grouped[t],
            marker_color=CORES_MAP[t], text=mensal_grouped[t],
            textposition='outside', textfont=dict(size=26, family="Arial Black"),
            legendgroup=t, showlegend=True
        ), row=4, col=1)

# ======================================================
# 5. ESTILIZAÇÃO FINAL
# ======================================================
# Desativa a legenda padrão
fig.update_layout(showlegend=False)

fig.update_layout(
    title=dict(
        text=f"<b>DASHBOARD QUALIDADE PRODUÇÃO - {data_input}</b>",
        x=0.5, y=0.98, font=dict(size=85, color=AZUL_PCP)
    ),
    template="plotly_white", 
    width=3508, height=2480,
    # Aumentamos a margem direita (r=600) para os gráficos terem respiro
    margin=dict(t=320, l=450, r=600, b=100), 
)

# --- COORDENADAS DA LEGENDA (FORA DA ÁREA DOS GRÁFICOS) ---
# x_base acima de 1.0 coloca a legenda na margem que criamos acima
x_base = 1.02   
y_topo = 0.50   
espaco_v = 0.025 

labels = ["Retrabalho", "Sucata", "Registro"]
cores = [AZUL_PCP, LARANJA_PCP, CINZA_REGISTRO]

# Título da Legenda
fig.add_annotation(
    x=x_base, y=y_topo + 0.04,
    text="<b>TRATATIVA:</b>",
    showarrow=False, xref="paper", yref="paper",
    font=dict(size=35, family="Arial Black", color=AZUL_PCP),
    xanchor="left"
)

for i, (label, cor) in enumerate(zip(labels, cores)):
    y_pos = y_topo - (i * espaco_v)
    
    # Desenha o Quadrado de Cor
    fig.add_shape(
        type="rect",
        xref="paper", yref="paper",
        x0=x_base, x1=x_base + 0.008,     
        y0=y_pos - 0.005, y1=y_pos + 0.005, 
        fillcolor=cor,
        line=dict(width=0),
    )
    
    # Texto da Legenda
    fig.add_annotation(
        x=x_base + 0.012, y=y_pos,
        text=f"<b>{label}</b>",
        showarrow=False, xref="paper", yref="paper",
        font=dict(size=28, family="Arial Black", color="#333"),
        xanchor="left"
    )

# --- RESTANTE DOS AJUSTES (EIXOS E ANOTAÇÕES) ---
# Filtramos as anotações para não resetar o tamanho da nossa legenda manual
for i in fig['layout']['annotations']:
    if "TRATATIVA" not in i['text'] and i['text'] not in ["<b>Retrabalho</b>", "<b>Sucata</b>", "<b>Registro</b>"]:
        i['font'] = dict(size=55, color=AZUL_PCP)

fig.update_yaxes(visible=False) 
fig.update_yaxes(visible=True, row=2, col=2, tickfont=dict(size=28, family="Arial Black", color=AZUL_PCP))
fig.update_xaxes(showticklabels=True, tickfont=dict(size=30, family="Arial Black"), row=3, col=1)
fig.update_xaxes(showticklabels=True, tickfont=dict(size=30, family="Arial Black"), row=3, col=2)
fig.update_xaxes(showticklabels=True, tickfont=dict(size=45, family="Arial Black"), row=4, col=1)
fig.update_xaxes(showticklabels=False, row=2, col=2)

arquivo = f"Relatorios/Dashboard_Qualidade.png"
if not os.path.exists("Relatorios"): os.makedirs("Relatorios")
fig.write_image(arquivo)
print(f"\n[OK] Dashboard gerado: {arquivo}")