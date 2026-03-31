import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
from datetime import datetime

# ==========================
#  CARREGAMENTO DE DADOS
# ==========================
try:
    dfs = pd.read_excel('Dados.xlsx', sheet_name=[''])
    df = dfs['']
except Exception as e:
    print(f"ERRO CRÍTICO: Não foi possível ler o Excel. Verifique se '' existe. Erro: {e}")
    exit()

# Limpeza e preparação
df.columns = df.columns.str.strip()
df["custo ma qualidade"] = df[["custo total retrabalho", "custo total refugos"]].sum(axis=1)

# Filtro do Mês
print("-" * 30)
mes_escolhido_input = input("Digite o mês para o relatório (Ex: janeiro): ").strip().lower()
if not mes_escolhido_input or mes_escolhido_input not in df['Mês'].str.lower().values:
    dados = df.iloc[0] 
    mes_final = str(dados["Mês"])
else:
    dados = df[df['Mês'].str.lower() == mes_escolhido_input].iloc[0]
    mes_final = mes_escolhido_input.capitalize()

total = dados["custo ma qualidade"]
meta = dados["Meta"]
cor_status = "#2E7D32" if total <= meta else "#D32F2F"

# ======================================================
#  LAYOUT IMPRESSÃO A4 (3508 x 2480 px)
# ======================================================

fig = make_subplots(
    rows=2, cols=2,
    specs=[[{"type": "indicator"}, {"type": "domain"}],
           [{"type": "xy", "colspan": 2}, None]],
    vertical_spacing=0.18,
    horizontal_spacing=0.1
)

#  DESEMPENHO FINANCEIRO
fig.add_trace(
    go.Indicator(
        mode = "number+delta+gauge",
        value = total,
        number = {'prefix': "R$ ", 'font': {'size': 110, 'color': '#003366'}, 'valueformat': ",.2f"},
        delta = {'reference': meta, 'position': "bottom", 'increasing': {'color': '#D32F2F'}, 'decreasing': {'color': '#2E7D32'}, 'valueformat': ",.2f"},
        gauge = {
            'axis': {'range': [None, max(total, meta) * 1.2], 'tickfont': {'size': 25}},
            'bar': {'color': cor_status, 'thickness': 0.6},
            'bgcolor': "white",
            'steps': [
                {'range': [0, meta], 'color': "rgba(46, 125, 50, 0.15)"},    
                {'range': [meta, max(total, meta) * 1.2], 'color': "rgba(211, 47, 47, 0.15)"} 
            ],
            'threshold': {'line': {'color': "black", 'width': 6}, 'thickness': 0.8, 'value': meta}
        },
        title = {'text': f"<b>DESEMPENHO FINANCEIRO - {mes_final.upper()}</b><br><span style='font-size:0.5em;color:gray'>Meta Máxima: R$ {meta:,.2f}</span>", 'font': {'size': 35}}
    ),
    row=1, col=1
)

# COMPOSIÇÃO DO CUSTO 
fig.add_trace(
    go.Pie(
        labels=["Retrabalho", "Refugos"],
        values=[dados["custo total retrabalho"], dados["custo total refugos"]],
        hole=0.6,
        marker=dict(colors=["#003366", "#FF6600"]),
        textinfo="label+percent",
        textposition="outside",
        textfont=dict(size=35, family="Arial Black"),
        showlegend=False,
        title={
            'text': "<b>COMPOSIÇÃO DO CUSTO</b><br><br><br>", # <br> adicionais criam o espaço solicitado
            'font': {'size': 35}, 
            'position': 'top center'
        }
    ),
    row=1, col=2
)

# EVOLUÇÃO MENSAL
fig.add_trace(
    go.Bar(
        x=df["Mês"], 
        y=df["custo ma qualidade"],
        name="Custo Real",
        marker_color='#003366',
        text=[f"R$ {v:,.0f}" for v in df["custo ma qualidade"]],
        textposition="outside",
        textfont=dict(size=25, family="Arial Black"),
        showlegend=True
    ),
    row=2, col=1
)

fig.add_trace(
    go.Scatter(
        x=df["Mês"], 
        y=df["Meta"],
        name="Meta (Limite)",
        mode='lines',
        line=dict(color='#D32F2F', width=6, dash='dot'),
        showlegend=True
    ),
    row=2, col=1
)

# AJUSTES FINAIS DE TÍTULOS E LEGENDA
fig.update_layout(
    title=dict(text="<b>RELATÓRIO QUALIDADE 2026</b>", x=0.5, y=0.98, font=dict(size=60, color='#003366')),
    paper_bgcolor="white",
    plot_bgcolor="white",
    width=3508, height=2480, 
    margin=dict(t=300, b=300, l=150, r=150),
    legend=dict(
        orientation="h",
        yanchor="bottom", y=-0.12,
        xanchor="center", x=0.5,
        font=dict(size=40, family="Arial Black"),
        bordercolor="Black", borderwidth=2
    )
)

fig.add_annotation(
    text="<b>GRÁFICO DE EVOLUÇÃO MENSAL (REAL VS META)</b>",
    xref="paper", yref="paper", x=0.5, y=0.48,
    showarrow=False, font=dict(size=35)
)

fig.update_xaxes(tickfont=dict(size=30, family="Arial Black"), row=2, col=1)
fig.update_yaxes(title="Valor (R$)", title_font=dict(size=30), tickfont=dict(size=25), row=2, col=1, gridcolor="#EEE")

# SALVAR IMAGEM
if not os.path.exists("Relatorios"): os.makedirs("Relatorios")
nome_arquivo = f"Relatorios/Relatorio_Qualidade_{mes_final}.png"
fig.write_image(nome_arquivo)

print(f"Sucesso! O relatório foi salvo em: {nome_arquivo}")

# --- GERAR LOG DE HISTÓRICO CORRIGIDO ---
tipo_script = "Impressão A4"
timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
log_texto = f"[{timestamp}] Relatório de {mes_final} gerado com sucesso! (Tipo: {tipo_script})\n"

with open("historico_relatorios.txt", "a", encoding="utf-8") as f:
    f.write(log_texto)

print(f"Histórico atualizado em 'historico_relatorios.txt'")
