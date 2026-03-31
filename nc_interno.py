import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
import numpy as np

# ======================================================
# 1. CONFIGURAÇÕES E LEITURA
# ======================================================
CAMINHO_EXCEL = 'Dados.xlsx'
ABA_NC = 'nc_interno'
DIRETORIO_SAIDA = "Relatorios"

AZUL_PCP = "#002D5F"
LARANJA_PCP = "#FE4E10"

try:
    df = pd.read_excel(CAMINHO_EXCEL, sheet_name=ABA_NC)
except Exception as e:
    print(f"Erro ao carregar Dados.xlsx: {e}")
    exit()

df.columns = df.columns.str.strip()
mapa = {"Mês": "mes", "total rodas produzidas": "total_prod", "nc interno": "nc_interno", "% retrabalho": "retrabalho", "meta": "meta"}
df = df.rename(columns=mapa).fillna(0)

# Conversão numérica
for col in ["total_prod", "nc_interno", "meta", "retrabalho"]:
    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

# Ajuste de escala (0.05 -> 5.0%)
if df["meta"].max() <= 1: df["meta"] *= 100
if df["retrabalho"].max() <= 1: df["retrabalho"] *= 100

# Preparação de dados para o gráfico (evita linhas descendo para zero em meses futuros)
df_plot = df.copy()
df_plot.loc[df_plot['total_prod'] == 0, ['total_prod', 'nc_interno', 'retrabalho']] = np.nan

# --- MAPEAMENTO DE MESES ABREVIADOS ---
mapeamento_meses = {
    'JANEIRO': 'JAN', 'FEVEREIRO': 'FEV', 'MARÇO': 'MAR', 'ABRIL': 'ABR',
    'MAIO': 'MAI', 'JUNHO': 'JUN', 'JULHO': 'JUL', 'AGOSTO': 'AGO',
    'SETEMBRO': 'SET', 'OUTUBRO': 'OUT', 'NOVEMBRO': 'NOV', 'DEZEMBRO': 'DEZ'
}
df_plot['mes_curto'] = df_plot['mes'].str.upper().map(mapeamento_meses)

# ======================================================
# 2. SELEÇÃO DO MÊS (FILTRO DINÂMICO)
# ======================================================
print("-" * 40)
mes_input = input("Mês para o Relatório (Ex: Janeiro ou ENTER para último): ").strip().upper()

if mes_input == "":
    linha = df[df['total_prod'] > 0].iloc[-1]
else:
    filtro = df[df["mes"].str.upper() == mes_input]
    linha = filtro.iloc[0] if not filtro.empty else df[df['total_prod'] > 0].iloc[-1]

mes_sel = linha["mes"]
retrabalho_sel = linha["retrabalho"]
meta_sel = linha["meta"]

# ======================================================
# 3. ESTRUTURA DO DASHBOARD (A3)
# ======================================================
fig = make_subplots(
    rows=2, cols=2,
    specs=[[{"colspan": 2}, None], [{"secondary_y": True}, {"type": "indicator"}]],
    column_widths=[0.75, 0.25], # Aumenta a largura do gráfico de evolução (75%)
    row_heights=[0.5, 0.5],
    vertical_spacing=0.2,
    horizontal_spacing=0.03,
    subplot_titles=(f"<b>% NC Interna vs Meta</b>", f"<b>Produção Total vs NC Interna</b>", "")
)

# --- 1. GRÁFICO SUPERIOR: % RETRABALHO VS META ---
fig.add_trace(
    go.Scatter(x=df_plot["mes_curto"], y=df_plot["retrabalho"], name="% Retrabalho",
               line=dict(color=LARANJA_PCP, width=12), marker=dict(size=22),
               mode="lines+markers+text", 
               text=[f"{v:.2f}%" if pd.notnull(v) else "" for v in df_plot["retrabalho"]],
               textposition="top center", 
               textfont=dict(size=38, family="Arial Black", color=LARANJA_PCP)),
    row=1, col=1
)

fig.add_trace(
    go.Scatter(x=df_plot["mes_curto"], y=df_plot["meta"], name="Meta",
               line=dict(color=AZUL_PCP, width=6, dash='dash'),
               mode="lines+text",
               text=[f"{v:.2f}%" for v in df_plot["meta"]],
               textposition="top center",
               textfont=dict(size=32, family="Arial Black", color=AZUL_PCP)),
    row=1, col=1
)

# --- 2. GRÁFICO INFERIOR: PRODUÇÃO VS NC (ESTENDIDO) ---
# Produção (Azul) - Eixo Principal
fig.add_trace(
    go.Scatter(x=df_plot["mes_curto"], y=df_plot["total_prod"],
               line=dict(color=AZUL_PCP, width=12), marker=dict(size=20),
               mode="lines+markers+text",
               text=[f"{int(v)}" if pd.notnull(v) else "" for v in df_plot["total_prod"]],
               textposition="top center", 
               textfont=dict(size=38, family="Arial Black", color=AZUL_PCP),
               name="Produção Total"),
    row=2, col=1, secondary_y=False
)

# NC Interna (Laranja) - Eixo Secundário
fig.add_trace(
    go.Scatter(x=df_plot["mes_curto"], y=df_plot["nc_interno"],
               line=dict(color=LARANJA_PCP, width=10, dash="dot"), marker=dict(size=18),
               mode="lines+markers+text",
               text=[f"{int(v)}" if pd.notnull(v) else "" for v in df_plot["nc_interno"]], 
               textposition="top center", 
               textfont=dict(size=42, family="Arial Black", color=LARANJA_PCP),
               name="NC Interna"),
    row=2, col=1, secondary_y=True
)

# --- 3. GAUGE (FIXO LARANJA) ---
fig.add_trace(
    go.Indicator(
        mode="gauge+number", value=retrabalho_sel,
        title={'text': f"<b>META:{meta_sel}% {mes_sel.upper()}</b>", 
               'font': {'size': 55, 'color': AZUL_PCP, 'family': "Arial Black"}},
        number={'suffix': "%", 'font': {'size': 100, 'color': LARANJA_PCP}, 'valueformat': '.2f'},
        gauge={'axis': {'range': [0, 10], 'tickfont': {'size': 40}},
               'bar': {'color': LARANJA_PCP},
               'threshold': {'line': {'color': "black", 'width': 12}, 'thickness': 0.8, 'value': meta_sel}}
    ), row=2, col=2
)

# ======================================================
# 4. AJUSTES FINAIS DE LAYOUT
# ======================================================
fig.update_layout(
    title=dict(text=f"<b>QUALIDADE 2026</b>",
               x=0.5, y=0.97, font=dict(size=90, color=AZUL_PCP, family="Arial Black")),
    template="plotly_white", width=3508, height=2480,
    margin=dict(t=320, l=120, r=120, b=150),
    showlegend=False
)

# Eixos X: Horizontal e Fonte ajustada
fig.update_xaxes(tickfont=dict(size=35, family="Arial Black", color=AZUL_PCP), tickangle=0)

# Eixos Y: Invisíveis para manter limpeza visual
fig.update_yaxes(visible=False)

# Logica de "Empilhamento": Azul em cima (*1.8) e Laranja embaixo (*6)
fig.update_yaxes(range=[0, df_plot['total_prod'].max(skipna=True)*1.8], row=2, col=1, secondary_y=False)
fig.update_yaxes(range=[0, df_plot['nc_interno'].max(skipna=True)*6], row=2, col=1, secondary_y=True)

# Fonte dos títulos dos subplots
for i in fig['layout']['annotations']:
    i['font'] = dict(size=60, color=AZUL_PCP, family="Arial Black")

# Salvar Arquivo
if not os.path.exists(DIRETORIO_SAIDA): os.makedirs(DIRETORIO_SAIDA)
arquivo = f"{DIRETORIO_SAIDA}/Dashboard_Final_{mes_sel}_2026.png"
fig.write_image(arquivo)

print(f"\n[OK] Dashboard gerado com sucesso: {arquivo}")