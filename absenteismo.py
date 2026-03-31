import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
from datetime import datetime

# ================================
# 1. LEITURA E TRATAMENTO DE DADOS
# ================================
try:
    dfs = pd.read_excel('Dados.xlsx', sheet_name=['absenteismo'])
    df = dfs['absenteismo']
except Exception as e:
    print(f"Erro ao carregar Dados.xlsx: {e}")
    exit()

df.columns = df.columns.str.strip()
df = df.rename(columns={
    "total horas falta": "Horas_Falta",
    "total horas trabalhadas": "Horas_Trabalhadas",
    "meta": "Meta"
})

df["Horas_Falta"] = pd.to_numeric(df["Horas_Falta"], errors="coerce").fillna(0)
df["Horas_Trabalhadas"] = pd.to_numeric(df["Horas_Trabalhadas"], errors="coerce").fillna(0)
df["Meta"] = pd.to_numeric(df["Meta"], errors="coerce").fillna(0)

if df["Meta"].max() <= 1:
    df["Meta"] = df["Meta"] * 100

df["Mês"] = df["Mês"].astype(str).str.strip()
df["Abs_%"] = (df["Horas_Falta"] / df["Horas_Trabalhadas"] * 100).fillna(0)

# ================================
# 2. DEFINIÇÕES DE CORES
# ================================
AZUL_ESCURO = "#002D5F"
LARANJA_ALERTA = "#FE4E10"
VERDE_SOFT = "rgba(40, 167, 69, 0.3)"
VERMELHO_SOFT = "rgba(254, 78, 16, 0.2)"

# ================================
# 3. ESCOLHA DO MÊS
# ================================
mes_input = input("Digite o mês ou ENTER para o último: ").strip().upper()
if mes_input == "":
    linha = df.iloc[-1]
else:
    filtro = df[df["Mês"].str.upper() == mes_input]
    if filtro.empty:
        print("Mês não encontrado."); exit()
    linha = filtro.iloc[0]

mes = linha["Mês"]
h_trabalhadas = linha["Horas_Trabalhadas"]
h_falta = linha["Horas_Falta"]
meta_valor = linha["Meta"]
abs_mes = linha["Abs_%"]

# ================================
# 4. CONSTRUÇÃO DO DASHBOARD 
# ================================
# Agora com apenas 2 colunas na linha de baixo
fig = make_subplots(
    rows=2, cols=2,
    specs=[
        [{"colspan": 2}, None],
        [{"type": "bar"}, {"type": "indicator"}]
    ],
    vertical_spacing=0.18,
    row_heights=[0.5, 0.5]
)

# --- A. GRÁFICO DE EVOLUÇÃO ---
fig.add_trace(
    go.Scatter(
        x=df["Mês"], y=df["Abs_%"],
        mode="lines+markers+text",
        line=dict(color=AZUL_ESCURO, width=12),
        marker=dict(size=24),
        text=[f"{v:.1f}%" if v > 0 else "" for v in df["Abs_%"]],
        textposition="top center",
        textfont=dict(size=45, family="Arial Black")
    ), row=1, col=1
)

# Linha de Meta
fig.add_trace(
    go.Scatter(
        x=df["Mês"], y=df["Meta"],
        mode="lines+text",
        line=dict(color=LARANJA_ALERTA, width=6, dash="dash"),
        text=[""] * (len(df)-1) + [f"Meta {meta_valor:.1f}%"],
        textposition="top right",
        textfont=dict(size=45, family="Arial Black", color=LARANJA_ALERTA),
        showlegend=False
    ), row=1, col=1
)

# --- B. GRÁFICO DE BARRAS (Esquerda) ---
fig.add_trace(
    go.Bar(
        x=["TRABALHADAS", "FALTAS"],
        y=[h_trabalhadas, h_falta],
        marker_color=[AZUL_ESCURO, LARANJA_ALERTA],
        text=[f"{int(h_trabalhadas)}h", f"{int(h_falta)}h"],
        textposition="outside",
        textfont=dict(size=70, family="Arial Black")
    ), row=2, col=1
)

# --- C. GRÁFICO GAUGE (Direita) ---
fig.add_trace(
    go.Indicator(
        mode="gauge+number",
        value=abs_mes,
        number={'suffix': "%", 'font': {'size': 130, 'family': 'Arial Black'}, 'valueformat': '.2f'},
        gauge={
            'axis': {'range': [0, 10], 'tickfont': {'size': 35}},
            'bar': {'color': AZUL_ESCURO},
            'threshold': {'line': {'color': "black", 'width': 8}, 'thickness': 0.8, 'value': meta_valor},
            'steps': [
                {'range': [0, meta_valor], 'color': VERDE_SOFT},
                {'range': [meta_valor, 10], 'color': VERMELHO_SOFT}
            ]
        },
        title={'text': f"STATUS {mes.upper()}", 'font': {'size': 60, 'family': 'Arial Black'}},
    ), row=2, col=2
)

# ================================
# 5. AJUSTES DE LAYOUT E ANOTAÇÕES
# ================================
fig.update_layout(
    title=dict(
        text=f"<b>PERFORMANCE DE ABSENTEÍSMO 2026</b>",
        x=0.5, y=0.96, font=dict(size=85)
    ),
    template="plotly_white",
    width=3508, height=2480,
    margin=dict(t=250, l=150, r=200, b=150),
    showlegend=False
)

# Estilização de Eixos
fig.update_xaxes(tickfont=dict(size=48, family="Arial Black"), row=1, col=1)
fig.update_xaxes(tickfont=dict(size=55, family="Arial Black"), row=2, col=1)
fig.update_yaxes(tickfont=dict(size=40), row=1, col=1, range=[0, 12])
fig.update_yaxes(visible=False, row=2, col=1)

# ANOTAÇÃO: Ajustada para y=0.44 para não encostar no eixo X
fig.add_annotation(
    x=0.08, y=0.44,
    xref="paper", yref="paper",
    text="▼ <i>Quanto menor o valor, melhor o desempenho</i>",
    showarrow=False,
    font=dict(size=42, color=LARANJA_ALERTA)
)

# ================================
# 6. SALVAMENTO
# ================================
if not os.path.exists("Relatorios"):
    os.makedirs("Relatorios")

arquivo_final = f"Relatorios/Dashboard_RH_{mes}_Final_V5.png"
fig.write_image(arquivo_final)

print(f"Sucesso! Relatório gerado: {arquivo_final}")
