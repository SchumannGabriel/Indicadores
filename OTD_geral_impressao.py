import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
from datetime import datetime

# ================================
# LEITURA DOS DADOS
# ================================
try:
    dfs = pd.read_excel('Dados.xlsx', sheet_name=['on-time-delivery OTD'])
    df = dfs['on-time-delivery OTD']
except Exception as e:
    print(f"Erro ao carregar Dados.xlsx: {e}")
    exit()

df.columns = df.columns.str.strip()
df = df.rename(columns={
    "quantidade entregue no periodo": "Total",
    "quantidade entregue em dia": "No_Prazo",
    "otd": "OTD",
    "meta": "Meta"
})

df["Total"] = pd.to_numeric(df["Total"], errors="coerce").fillna(0)
df["No_Prazo"] = pd.to_numeric(df["No_Prazo"], errors="coerce").fillna(0)
df["Meta"] = pd.to_numeric(df["Meta"], errors="coerce").fillna(0)
df["Meta"] = df["Meta"] * 100
df["Mês"] = df["Mês"].astype(str).str.strip()
df["OTD_%"] = (df["No_Prazo"] / df["Total"] * 100).fillna(0)

# ================================
# CORES
# ================================
AZUL = "#002D5F"
LARANJA = "#FE4E10"
VERMELHO = "#FF0000"

# ================================
# ESCOLHA DO MÊS
# ================================
print("-" * 40)
mes_input = input("Digite o mês ou ENTER para último mês: ").strip().upper()
linha = df.iloc[-1] if mes_input == "" else df[df["Mês"].str.upper() == mes_input].iloc[0]

mes = linha["Mês"]
total = linha["Total"]
no_prazo = linha["No_Prazo"]
meta_valor = linha["Meta"]
otd_mes = linha["OTD_%"]
atraso = total - no_prazo

# ================================
# DASHBOARD - AJUSTE DE ESTRUTURA
# ================================
fig = make_subplots(
    rows=3, cols=2,
    specs=[
        [{"type": "indicator", "colspan": 2}, None],
        [{"colspan": 2}, None],
        [{"type": "bar"}, {"type": "domain"}]
    ],
    vertical_spacing=0.08, # Espaço para a legenda não encavalar
    row_heights=[0.20, 0.35, 0.45] 
)

# 1. INDICADOR (Ajustado para não bater no título)
fig.add_trace(
    go.Indicator(
        mode="number",
        value=otd_mes,
        number={'suffix': "%", 'font': {'size': 150, 'family': 'Arial Black', 'color': AZUL}, 'valueformat': '.1f'},
        title={'text': f"OTD {mes.upper()}", 'font': {'size': 45, 'family': 'Arial Black'}, 'align': 'center'},
        domain={'y': [0, 0.8]} # Limita o indicador para baixo do título
    ), row=1, col=1
)

# 2. EVOLUÇÃO OTD
fig.add_trace(
    go.Scatter(
        x=df["Mês"], y=df["OTD_%"],
        mode="lines+markers+text",
        line=dict(color=AZUL, width=12),
        marker=dict(size=22),
        text=[f"{v:.0f}%" for v in df["OTD_%"]],
        textposition="top center",
        textfont=dict(size=45, family="Arial Black")
    ), row=2, col=1
)

# Linha Meta
fig.add_trace(
    go.Scatter(
        x=df["Mês"], y=df["Meta"],
        mode="lines+text",
        line=dict(color=VERMELHO, width=6, dash="dash"),
        text=[""] * (len(df) - 1) + [f" META {meta_valor:.0f}%"],
        textposition="middle left",
        textfont=dict(size=45, family="Arial Black", color=VERMELHO)
    ), row=2, col=1
)

# 3. BARRAS E DONUT
fig.add_trace(
    go.Bar(
        x=["TOTAL ENTREGUE", "NO PRAZO"],
        y=[total, no_prazo],
        marker_color=[AZUL, LARANJA],
        text=[f"{int(total)}", f"{int(no_prazo)}"],
        textposition="outside",
        cliponaxis=False,
        textfont=dict(size=90, family="Arial Black")
    ), row=3, col=1
)

fig.add_trace(
    go.Pie(
        labels=["No Prazo", "Atraso"],
        values=[no_prazo, atraso],
        hole=0.6,
        marker=dict(colors=[AZUL, LARANJA]),
        texttemplate="<b>%{label}</b><br>%{percent}",
        textposition="outside",
        textfont=dict(size=50, family="Arial Black")
    ), row=3, col=2
)

# ================================
# LAYOUT E MARGENS
# ================================
fig.update_layout(
    title=dict(
        text=f"<b>PERFORMANCE ON TIME DELIVERY 2026</b>",
        x=0.5, y=0.98,
        yanchor='top',
        font=dict(size=90, family="Arial Black", color=AZUL)
    ),
    template="plotly_white",
    width=3508, height=2480,
    margin=dict(t=120, l=100, r=280, b=100), 
    showlegend=False
)

fig.update_xaxes(tickfont=dict(size=48, family="Arial Black"), row=2, col=1)
fig.update_yaxes(tickfont=dict(size=38), range=[0, 125], row=2, col=1)
fig.update_xaxes(tickfont=dict(size=50, family="Arial Black"), row=3, col=1)
fig.update_yaxes(visible=False, row=3, col=1)

# LEGENDA ENCIMA DO GRÁFICO DE LINHA
fig.add_annotation(
    x=0, y=0.74, # Posicionada logo acima do gráfico de evolução
    xref="paper", yref="paper",
    text="▲ <i>Quanto maior o valor, melhor o desempenho</i>",
    showarrow=False,
    font=dict(size=45, color=AZUL),
    align="left"
)

# ================================
# SALVAR
# ================================
if not os.path.exists("Relatorios"): os.makedirs("Relatorios")
arquivo = f"Relatorios/Dashboard_PCP_{mes}_Corrigido.png"
fig.write_image(arquivo)
print(f"Dashboard gerado: {arquivo}")