#Deixa off por enquanto

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
from datetime import datetime

# ======================================================
# 1. CONFIGURAÇÕES E LEITURA
# ======================================================
CAMINHO_EXCEL = 'Dados.xlsx'
ABA_RH = 'custo_funcionario'
DIRETORIO_SAIDA = "Relatorios"

AZUL_PCP = "#002D5F"    
LARANJA_PCP = "#FE4E10" 
VERMELHO_ALERTA = "#FF0000"

try:
    df = pd.read_excel(CAMINHO_EXCEL, sheet_name=ABA_RH)
except Exception as e:
    print(f"Erro ao carregar aba {ABA_RH}: {e}")
    exit()

df.columns = df.columns.str.strip().str.lower()
mapa = {'mes': 'mes', 'total folha': 'folha', 'total faturamento': 'faturamento', 'meta': 'meta'}
df = df.rename(columns=mapa).fillna(0)

# Cálculos e Ajuste de Meta
df['perc_custo'] = (df['folha'] / df['faturamento'] * 100).fillna(0)
df['meta'] = df['meta'].apply(lambda x: x * 100 if x <= 1 else x)

# Criamos uma versão para a linha de evolução que substitui 0 por None
# Isso faz a linha "parar" no último mês preenchido, mas mantém o eixo X completo
df_plot = df.copy()
df_plot.loc[df_plot['faturamento'] == 0, 'perc_custo'] = None

# ======================================================
# 2. SELEÇÃO DO MÊS
# ======================================================
print("-" * 40)
mes_input = input("Mês do KPI Financeiro (ENTER para último): ").strip().upper()

if mes_input == "":
    # Pega o último mês que tem faturamento real
    linha = df[df['faturamento'] > 0].iloc[-1]
else:
    filtro = df[df["mes"].str.upper() == mes_input]
    if filtro.empty:
        print("Mês não encontrado.")
        exit()
    linha = filtro.iloc[0]

mes = linha["mes"]
folha = linha["folha"]
faturamento = linha["faturamento"]
perc_real = linha["perc_custo"]
meta_limite = linha["meta"]

cor_status = VERMELHO_ALERTA if perc_real > meta_limite else AZUL_PCP

# ======================================================
# 3. DASHBOARD (ESTRUTURA A3 - 3508x2480)
# ======================================================
fig = make_subplots(
    rows=2, cols=2,
    specs=[[{"colspan": 2}, None], [{"type": "bar"}, {"type": "indicator"}]],
    row_heights=[0.5, 0.5],
    vertical_spacing=0.15,
    subplot_titles=("<b>Evolução Mensal do Índice (%)</b>", "<b>Comparativo de Valores (R$)</b>", "")
)

# --- 1. EVOLUÇÃO HISTÓRICA (Mantém os 12 meses no eixo, mas linha para no dado real) ---
fig.add_trace(
    go.Scatter(x=df_plot["mes"], y=df_plot["perc_custo"], 
               line=dict(color=AZUL_PCP, width=12), marker=dict(size=20),
               mode="lines+markers+text", 
               text=[f"{v:.1f}%" if v is not None else "" for v in df_plot["perc_custo"]],
               textposition="top center", textfont=dict(size=45, family="Arial Black"),
               connectgaps=False), # Importante: não conecta com o vazio
    row=1, col=1
)

# --- 2. BARRAS ---
fig.add_trace(
    go.Bar(x=["FATURAMENTO", "FOLHA"], y=[faturamento, folha],
           marker_color=[AZUL_PCP, LARANJA_PCP],
           text=[f"R$ {faturamento:,.0f}", f"R$ {folha:,.0f}"],
           textposition="outside", textfont=dict(size=85, family="Arial Black")),
    row=2, col=1
)

# --- 3. GAUGE (EFICIÊNCIA) ---
fig.add_trace(
    go.Indicator(
        mode="gauge+number", value=perc_real,
        title={'text': f"<b>Eficiência</b><br><span style='font-size:50px'>Meta: {meta_limite:.0f}%</span>", 'font': {'size': 60}},
        number={'suffix': "%", 'font': {'size': 200, 'color': cor_status}, 'valueformat': '.1f'},
        gauge={'axis': {'range': [0, 50], 'tickfont': {'size': 45}},
               'bar': {'color': cor_status},
               'threshold': {'line': {'color': "black", 'width': 12}, 'thickness': 0.8, 'value': meta_limite}}
    ), row=2, col=2
)

# ======================================================
# 4. LAYOUT FINAL E IMPRESSÃO
# ======================================================
fig.update_layout(
    title=dict(text=f"<b>CUSTO FUNCIONÁRIO - {mes.upper()}</b>",
               x=0.5, font=dict(size=100, color=AZUL_PCP)),
    template="plotly_white", width=3508, height=2480,
    margin=dict(t=300, l=150, r=150, b=150), showlegend=False
)

# Formatação dos títulos dos gráficos
for i in fig['layout']['annotations']: i['font'] = dict(size=60, color="#333")

# Eixos
fig.update_xaxes(tickfont=dict(size=50, family="Arial Black"), row=1, col=1)
fig.update_yaxes(range=[0, max(df['perc_custo'].max(), 30)], tickfont=dict(size=40), row=1, col=1)
fig.update_xaxes(tickfont=dict(size=70, family="Arial Black"), row=2, col=1)
fig.update_yaxes(visible=False, row=2, col=1)

# ======================================================
# 5. SALVAMENTO
# ======================================================
if not os.path.exists(DIRETORIO_SAIDA): os.makedirs(DIRETORIO_SAIDA)
arquivo = f"{DIRETORIO_SAIDA}/Dashboard_Custo_Funcionario_{mes}_2026.png"
fig.write_image(arquivo)

print(f"\n[OK] Dashboard padronizado gerado: {arquivo}")

