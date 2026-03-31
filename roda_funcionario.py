import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
import numpy as np

# ======================================================
# 1. CONFIGURAÇÕES E LEITURA
# ======================================================
CAMINHO_EXCEL = 'Dados.xlsx'
ABA_PROD = 'roda_por_funcionario'
DIRETORIO_SAIDA = "Relatorios"

AZUL_PCP = "#002D5F"    
LARANJA_PCP = "#FE4E10" 

try:
    df = pd.read_excel(CAMINHO_EXCEL, sheet_name=ABA_PROD)
except Exception as e:
    print(f"Erro ao carregar aba {ABA_PROD}: {e}")
    exit()

df.columns = df.columns.str.strip()
mapa = {
    'Mês': 'mes',
    'total roda faturada': 'rodas',
    'N de funcionarios': 'func',
    'produtividade media por funcionario': 'prod',
    'meta': 'meta'
}
df = df.rename(columns=mapa).fillna(0)

df['meta'] = df['meta'].apply(lambda x: x * 100 if x <= 1 else x)
df_plot = df.copy()
# Mantemos os dados para garantir que o eixo funcione
df_plot.loc[df_plot['rodas'] == 0, ['prod', 'rodas', 'func']] = np.nan

# ======================================================
# 2. SELEÇÃO DO MÊS
# ======================================================
print("-" * 40)
mes_input = input("Mês da Produtividade (ENTER para último): ").strip().upper()
if mes_input == "":
    linha = df[df['rodas'] > 0].iloc[-1]
else:
    filtro = df[df["mes"].str.upper() == mes_input]
    linha = filtro.iloc[0] if not filtro.empty else df[df['rodas'] > 0].iloc[-1]

mes = linha["mes"]
meta_valor = linha["meta"]

# ======================================================
# 3. DASHBOARD (COM EIXO SECUNDÁRIO NO GRÁFICO 2)
# ======================================================
fig = make_subplots(
    rows=2, cols=1,
    row_heights=[0.5, 0.5],
    vertical_spacing=0.25,
    # Ativando eixo secundário apenas para a linha de funcionários
    specs=[[{"secondary_y": False}], [{"secondary_y": True}]],
    subplot_titles=(f"<b>Evolução da Produtividade Mensal</b>", 
                    f"<b>Comparativo: Volume de Rodas vs Funcionários</b>")
)

# --- 1. GRÁFICO SUPERIOR (Produtividade) ---
fig.add_trace(
    go.Scatter(x=df_plot["mes"], y=df_plot["prod"], 
               line=dict(color=LARANJA_PCP, width=12), marker=dict(size=22),
               mode="lines+markers+text", 
               text=[f"{int(v)}" if pd.notnull(v) else "" for v in df_plot["prod"]],
               textposition="top center", 
               textfont=dict(size=40, family="Arial Black", color=LARANJA_PCP), 
               name="Produtividade", showlegend=False),
    row=1, col=1
)

fig.add_trace(
    go.Scatter(x=df_plot["mes"], y=df_plot["meta"],
               mode="lines+text",
               line=dict(color=AZUL_PCP, width=6, dash="dash"),
               text=[""] * (len(df_plot)-1) + [f" META: {int(meta_valor)} "],
               textposition="top left",
               textfont=dict(size=45, family="Arial Black", color=AZUL_PCP),
               name="Meta", showlegend=False),
    row=1, col=1
)

# --- 2. GRÁFICO INFERIOR (Rodas vs Funcionários) ---

# Rodas (Azul)
fig.add_trace(
    go.Scatter(x=df_plot["mes"], y=df_plot["rodas"],
               line=dict(color=AZUL_PCP, width=12), marker=dict(size=20),
               mode="lines+markers+text",
               text=[f"{int(v)}" if pd.notnull(v) else "" for v in df_plot["rodas"]],
               textposition="top center", # Valor das rodas em cima
               textfont=dict(size=40, family="Arial Black", color=AZUL_PCP),
               name="Rodas Totais", showlegend=False),
    row=2, col=1, secondary_y=False
)

# Funcionários (Laranja)
fig.add_trace(
    go.Scatter(x=df_plot["mes"], y=df_plot["func"],
               line=dict(color=LARANJA_PCP, width=10, dash="dot"), marker=dict(size=18),
               mode="lines+markers+text",
               text=[f"{int(v)}" if pd.notnull(v) else "" for v in df_plot["func"]], 
               textposition="top center", # Mudado para 'top' para subir o número
               textfont=dict(size=45, family="Arial Black", color=LARANJA_PCP),
               name="Funcionários", showlegend=False),
    row=2, col=1, secondary_y=True
)

# --- AJUSTE DOS EIXOS (O segredo para a posição) ---

# Eixo das Rodas (Esquerda)
fig.update_yaxes(range=[0, df_plot['rodas'].max(skipna=True)*1.8], row=2, col=1, secondary_y=False)

# Eixo dos Funcionários (Direita) 
# Ajustamos o range para que o 'topo' da linha laranja fique logo abaixo da azul
fig.update_yaxes(range=[0, df_plot['func'].max(skipna=True)*6], # Aumentar o multiplicador (6) 'empurra' a linha para baixo
                 visible=False, row=2, col=1, secondary_y=True)

# ======================================================
# 4. LEGENDAS E ESTILIZAÇÃO
# ======================================================
fig.add_annotation(dict(x=0.98, y=1.08, xref="paper", yref="paper", showarrow=False,
                   text=f"<span style='color:{LARANJA_PCP}'>■</span> Produtividade  <span style='color:{AZUL_PCP}'>--</span> Meta",
                   font=dict(size=45, family="Arial Black"), align="right"))

fig.add_annotation(dict(x=0.98, y=0.45, xref="paper", yref="paper", showarrow=False,
                   text=f"<span style='color:{AZUL_PCP}'>■</span> Rodas Totais  <span style='color:{LARANJA_PCP}'>●</span> QTD Funcionários",
                   font=dict(size=45, family="Arial Black"), align="right"))

fig.update_layout(
    title=dict(text=f"<b>PRODUTIVIDADE POR FUNCIONÁRIO - {mes.upper()}</b>",
               x=0.5, y=0.97, font=dict(size=85, color=AZUL_PCP)),
    template="plotly_white", width=3508, height=2480,
    margin=dict(t=380, l=150, r=150, b=150),
    showlegend=False
)

# Ajuste dos Títulos dos Subplots
for i in fig['layout']['annotations'][:2]:
    i['font'] = dict(size=60, color=AZUL_PCP)

# --- AJUSTE DOS EIXOS PARA NÃO CORTAR NÚMEROS ---
fig.update_yaxes(range=[df_plot['prod'].min(skipna=True)*0.5, df_plot['prod'].max(skipna=True)*1.6], visible=False, row=1, col=1)

# Eixo das Rodas (Esquerda)
fig.update_yaxes(range=[0, df_plot['rodas'].max(skipna=True)*1.8], visible=False, row=2, col=1, secondary_y=False)

# Eixo dos Funcionários (Direita) - Essencial para o número aparecer!
fig.update_yaxes(range=[df_plot['func'].min(skipna=True)*0.1, df_plot['func'].max(skipna=True)*1.4], visible=False, row=2, col=1, secondary_y=True)

fig.update_xaxes(tickfont=dict(size=50, family="Arial Black", color=AZUL_PCP), row=1, col=1)
fig.update_xaxes(tickfont=dict(size=50, family="Arial Black", color=AZUL_PCP), row=2, col=1)

# Salvar
if not os.path.exists(DIRETORIO_SAIDA): os.makedirs(DIRETORIO_SAIDA)
arquivo = f"{DIRETORIO_SAIDA}/Dashboard_Produtividade_{mes}.png"
fig.write_image(arquivo)

print(f"\n[OK] Dashboard Gerado: {arquivo}")