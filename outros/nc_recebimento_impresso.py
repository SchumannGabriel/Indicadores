import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
from datetime import datetime

# ==========================
#  CARREGAMENTO DE DADOS
# ==========================
try:
    dfs = pd.read_excel('', sheet_name=[''])
    df = dfs['']
except Exception as e:
    print(f"ERRO CRÍTICO: Não foi possível ler a aba ''. Erro: {e}")
    exit()

df.columns = df.columns.str.strip()

col_mes = "Mês"
col_nc = "Total NC recebimento"
col_insp = "total inspeções"
col_meta = "meta"

# ==========================
# TRATAMENTO NUMÉRICO
# ==========================
df[col_nc] = pd.to_numeric(df[col_nc], errors='coerce').fillna(0)
df[col_insp] = pd.to_numeric(df[col_insp], errors='coerce').fillna(0)
df[col_meta] = pd.to_numeric(df[col_meta], errors='coerce').fillna(0)

# Corrige meta caso venha decimal estranho (ex: 0.001)
df[col_meta] = df[col_meta].round(0)

# ==========================
# ORDENAR MESES
# ==========================
ordem_meses = [
    "janeiro","fevereiro","março","abril","maio","junho",
    "julho","agosto","setembro","outubro","novembro","dezembro"
]

df[col_mes] = df[col_mes].str.lower()
df[col_mes] = pd.Categorical(df[col_mes], categories=ordem_meses, ordered=True)
df = df.sort_values(col_mes)

# ==========================
# CÁLCULO TAXA NC (%)
# ==========================
df["Taxa NC (%)"] = df.apply(
    lambda row: (row[col_nc] / row[col_insp] * 100) if row[col_insp] > 0 else 0,
    axis=1
)

# ==========================
# SELEÇÃO DO MÊS
# ==========================
print("-" * 30)
mes_escolhido_input = input("Digite o mês para o relatório de NC (Ex: janeiro): ").strip().lower()

if not mes_escolhido_input or mes_escolhido_input not in df[col_mes].astype(str).values:
    dados = df.iloc[-1]
else:
    dados = df[df[col_mes].astype(str) == mes_escolhido_input].iloc[0]

mes_final = str(dados[col_mes]).capitalize()

taxa_mes = (dados[col_nc] / dados[col_insp] * 100) if dados[col_insp] > 0 else 0

# Cor automática
cor_principal = "#D32F2F" if dados[col_nc] > dados[col_meta] else "#2E7D32"

ano_atual = datetime.now().year

# ==========================
# LAYOUT
# ==========================
fig = make_subplots(
    rows=2, cols=1,
    specs=[[{"type": "indicator"}],
           [{"type": "xy"}]],
    vertical_spacing=0.15,
    row_heights=[0.4, 0.6]
)

# ==========================
# INDICADOR PREMIUM (GAUGE)
# ==========================
valor_maximo_gauge = max(df[col_nc].max()*1.3, df[col_meta].max()*2)

fig.add_trace(
    go.Indicator(
        mode="gauge+number+delta",
        value=dados[col_nc],
        number={
            'font': {'size': 120, 'color': cor_principal}
        },
        delta={
            'reference': dados[col_meta],
            'relative': True,
            'valueformat': '.1%',
            'position': "bottom",
            'increasing': {'color': '#D32F2F'},
            'decreasing': {'color': '#2E7D32'}
        },
        title={
            'text': f"""
            <b>NC RECEBIMENTO</b><br>
            <span style='font-size:0.6em;color:#555'>
            {mes_final.upper()} | Meta: {int(dados[col_meta])} | 
            Inspeções: {int(dados[col_insp])} | 
            Taxa: {taxa_mes:.2f}%
            </span>
            """,
            'font': {'size': 45}
        },
        gauge={
            'axis': {'range': [0, valor_maximo_gauge]},
            'bar': {'color': cor_principal},
            'steps': [
                {'range': [0, dados[col_meta]], 'color': '#E8F5E9'},
                {'range': [dados[col_meta], valor_maximo_gauge], 'color': '#FFEBEE'}
            ],
            'threshold': {
                'line': {'color': "black", 'width': 4},
                'thickness': 0.8,
                'value': dados[col_meta]
            }
        }
    ),
    row=1, col=1
)

# ==========================
# GRÁFICO INFERIOR
# ==========================
fig.add_trace(
    go.Bar(
        x=df[col_mes],
        y=df[col_insp],
        name="Total Inspeções",
        marker_color='#B0BEC5',
        text=df[col_insp],
        textposition="outside",
        textfont=dict(size=16)
    ),
    row=2, col=1
)

fig.add_trace(
    go.Scatter(
        x=df[col_mes],
        y=df[col_nc],
        name="Total NCs",
        mode='lines+markers+text',
        line=dict(color='#D32F2F', width=5),
        marker=dict(size=10),
        text=df[col_nc],
        textposition="top center",
        textfont=dict(size=18, color="#D32F2F")
    ),
    row=2, col=1
)

fig.add_trace(
    go.Scatter(
        x=df[col_mes],
        y=df[col_meta],
        name="Meta NC",
        mode='lines',
        line=dict(color='black', width=3, dash='dot')
    ),
    row=2, col=1
)

# ==========================
# AJUSTES FINAIS
# ==========================
fig.update_layout(
    title=dict(
        text=f"<b>QUALIDADE - NC RECEBIMENTO {ano_atual}</b>",
        x=0.5,
        y=0.97,
        font=dict(size=55, color='#003366')
    ),
    paper_bgcolor="white",
    plot_bgcolor="white",
    font=dict(family="Arial Black"),
    width=3508, height=2480,
    margin=dict(t=350, b=250, l=150, r=150),
    legend=dict(
        orientation="h",
        yanchor="bottom", y=-0.1,
        xanchor="center", x=0.5,
        font=dict(size=25),
        bordercolor="Black",
        borderwidth=2
    )
)

fig.update_xaxes(tickfont=dict(size=20), row=2, col=1)
fig.update_yaxes(title="Quantidade",
                 title_font=dict(size=25),
                 tickfont=dict(size=18),
                 gridcolor="#EEE",
                 row=2, col=1)

# ==========================
# SALVAR
# ==========================
if not os.path.exists("Relatorios"):
    os.makedirs("Relatorios")

nome_arquivo = f"Relatorios/Relatorio_NC_Recebimento_{mes_final}.png"
fig.write_image(nome_arquivo)

print(f"Relatório executivo salvo em: {nome_arquivo}")

# --- GERAR LOG DE HISTÓRICO---
tipo_script = "Impressão A4"
timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
log_texto = f"[{timestamp}] Relatório de {mes_final} gerado com sucesso! (Tipo: {tipo_script})\n"

with open("historico_relatorios.txt", "a", encoding="utf-8") as f:
    f.write(log_texto)

print(f"Histórico atualizado em 'historico_relatorios.txt'")
