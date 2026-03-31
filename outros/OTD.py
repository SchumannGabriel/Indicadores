import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os   
from datetime import datetime

# ==========================
# DADOS
# ==========================
try:
    dfs = pd.read_excel('Dados.xlsx', sheet_name=['on-time-delivery OTD'])
    df = dfs['on-time-delivery OTD']
except:
    # Dados de exemplo
    df = pd.DataFrame({"Mês": ["Janeiro"], "Total": [6843], "No_Prazo": [4344]})

df.columns = df.columns.str.strip()
df = df.rename(columns={
    "quantidade entregue no periodo": "Total",
    "quantidade entregue em dia": "No_Prazo"
})

df = df.dropna(subset=["Total", "Mês"])
df["Mês"] = df["Mês"].astype(str).str.strip()

# ==========================
# ESCOLHA DO MÊS
# ==========================
print("-" * 40)
mes_input = input("Digite o mês (ou ENTER para último mês): ").strip().lower()

if mes_input == "":
    df_mes = df.tail(1)
else:
    df_mes = df[df["Mês"].str.lower() == mes_input]
    if df_mes.empty:
        print("Mês não encontrado na planilha.")
        exit()

linha = df_mes.iloc[0]
mes_escolhido = linha["Mês"].upper()

total = float(linha["Total"])
no_prazo = float(linha["No_Prazo"])
atraso = max(0, total - no_prazo)

# ==========================
# CORES
# ==========================
COR_AZUL = "#002D5F"
COR_LARANJA = "#FE4E10"

# ==========================
# LAYOUT A4
# ==========================
fig = make_subplots(
    rows=1, cols=2,
    specs=[[{"type": "bar"}, {"type": "domain"}]],
    column_widths=[0.5, 0.5],
    horizontal_spacing=0.25
)

# BARRAS
fig.add_trace(
    go.Bar(
        x=["Total Entregas", "No Prazo"],
        y=[total, no_prazo],
        marker_color=[COR_AZUL, COR_LARANJA],
        width=0.5,
        text=[int(total), int(no_prazo)],
        textposition="outside",
        textfont=dict(size=45, color="black", family="Arial Black"),
        showlegend=False
    ),
    row=1, col=1
)

# ROSCA
fig.add_trace(
    go.Pie(
        labels=["No Prazo", "Atraso"],
        values=[no_prazo, atraso],
        hole=0.5,
        marker=dict(colors=[COR_AZUL, COR_LARANJA]),
        textinfo="label+value+percent",
        textposition="outside",
        textfont=dict(size=38, family="Arial Black"),
        showlegend=False,
        sort=False
    ),
    row=1, col=2
)

# LAYOUT FINAL A4
fig.update_layout(
    title=dict(
        text=f"<b>RELATÓRIO PCP - ON TIME DELIVERY (OTD)</b><br><sup>{mes_escolhido} 2026</sup>",
        x=0.5,
        y=0.95,
        font=dict(size=60, color="black", family="Arial")
    ),
    template="plotly_white",
    width=3508,
    height=2480,
    margin=dict(t=400, l=150, r=300, b=200),
)

fig.update_xaxes(
    tickfont=dict(size=40, color="black", family="Arial"),
    row=1, col=1
)

fig.update_yaxes(
    range=[0, total * 1.35],
    tickfont=dict(size=30),
    showgrid=True,
    gridcolor="#EEE",
    row=1, col=1
)

# ==========================
# SALVAR
# ==========================
if not os.path.exists("Relatorios"):
    os.makedirs("Relatorios")

nome_arquivo = f"OTD_{mes_escolhido}_2026.png"
fig.write_image(nome_arquivo)

print(f"Arquivo gerado: {nome_arquivo}")
