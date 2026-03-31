import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os

# ============================================
# CONFIGURAÇÕES DE IMPRESSÃO A3 
# ===========================================
CONFIG = {
    "Caminho_Excel": "Dados.xlsx",
    "Aba": "indice_devolucoes",
    "Diretorio_Saida": "Relatorios",
    "Cores": {
        "Primaria": "#002D5F",   # Azul Escuro
        "Destaque": "#FE4E10",   # Laranja
        "Sucesso": "#28a745",    # Verde
        "Erro": "#D62728",       # Vermelho
    },
    "DPI_Scale": 3,
    "Largura_Pixels": 1587,
    "Altura_Pixels": 1123
}

def carregar_e_processar():
    if not os.path.exists(CONFIG["Caminho_Excel"]):
        print(f"Erro: Arquivo {CONFIG['Caminho_Excel']} não encontrado.")
        return None, None

    try:
        df = pd.read_excel(CONFIG["Caminho_Excel"], sheet_name=CONFIG["Aba"])
        df.columns = df.columns.str.strip().str.lower()
        
        mapa = {
            'rodas grandes devolvidas': 'rogr_dev', 'rodas grandes entregues': 'rogr_ent',
            'rodas pequenas devolvidas': 'rope_dev', 'rodas pequenas entregues': 'rope_ent',
            'rodas medias devolvidas': 'rome_dev', 'rodas medias entregues': 'rome_ent',
            'roda pequena bipartida': 'ro_bipartida',
            'mês': 'mes', 'meta': 'meta'
        }
        df = df.rename(columns=mapa).fillna(0)
        df['mes'] = df['mes'].astype(str).str.strip()

        df['total_dev'] = df[['rogr_dev', 'rope_dev', 'rome_dev']].sum(axis=1)
        
        cols_saida = ['rogr_ent', 'rope_ent', 'rome_ent', 'ro_bipartida', 'kits', 'bacias', 'flanges', 'tiras de anel']
        presentes = [c for c in cols_saida if c in df.columns]
        
        df['total_saidas'] = df[presentes].sum(axis=1)
        df['perc_dev'] = (df['total_dev'] / df['total_saidas'] * 100).replace([float('inf')], 0).fillna(0)
        
        mapa_meses = {
            'JANEIRO': 'JAN', 'FEVEREIRO': 'FEV', 'MARÇO': 'MAR', 'ABRIL': 'ABR',
            'MAIO': 'MAI', 'JUNHO': 'JUN', 'JULHO': 'JUL', 'AGOSTO': 'AGO',
            'SETEMBRO': 'SET', 'OUTUBRO': 'OUT', 'NOVEMBRO': 'NOV', 'DEZEMBRO': 'DEZ'
        }
        df['mes_curto'] = df['mes'].str.upper().map(mapa_meses).fillna(df['mes'])
        
        return df, presentes

    except Exception as e:
        print(f"Erro no processamento: {e}")
        return None, None

def gerar_relatorio_a3(df, colunas_producao):
    meses_disponiveis = df['mes'].unique()
    print(f"\nMeses encontrados: {', '.join(meses_disponiveis)}")
    escolha = input("Escolha o mês (ou ENTER para o último): ").strip()

    if escolha == "" or escolha.upper() not in [m.upper() for m in meses_disponiveis]:
        linha_atual = df.iloc[-1]
    else:
        linha_atual = df[df['mes'].str.upper() == escolha.upper()].iloc[0]

    mes_nome = linha_atual['mes'].upper()
    meta_valor = 0.40 
    cor_laranja = CONFIG["Cores"]["Destaque"]

    fig = make_subplots(
        rows=3, cols=2,
        column_widths=[0.4, 0.6],
        row_heights=[0.2, 0.4, 0.4],
        vertical_spacing=0.12,
        horizontal_spacing=0.08,
        subplot_titles=(
            "<b>VOLUME TOTAL DE SAÍDAS</b>", "<b>VOLUME TOTAL DE DEVOLUÇÕES</b>",
            "<b>STATUS DO ÍNDICE (%)</b>", "<b>DEVOLUÇÕES POR CATEGORIA (UN)</b>",
            "<b>MIX DE PRODUÇÃO ATIVO</b>", "<b>EVOLUÇÃO HISTÓRICA DO ÍNDICE (%)</b>"
        ),
        specs=[[{"type": "indicator"}, {"type": "indicator"}],
               [{"type": "indicator"}, {"type": "bar"}],
               [{"type": "bar"}, {"type": "scatter"}]]
    )

    # --- 1. CARDS ---
    fig.add_trace(go.Indicator(
        mode="number", value=int(linha_atual['total_saidas']),
        number={'font': {'size': 80, 'color': CONFIG["Cores"]["Primaria"]}, 'valueformat': 'd'}
    ), row=1, col=1)

    fig.add_trace(go.Indicator(
        mode="number", value=int(linha_atual['total_dev']),
        number={'font': {'size': 80, 'color': CONFIG["Cores"]["Erro"]}, 'valueformat': 'd'}
    ), row=1, col=2)

    # --- 2. GAUGE ---
    fig.add_trace(go.Indicator(
        mode="gauge+number", value=linha_atual['perc_dev'],
        number={'suffix': "%", 'font': {'size': 55, 'color': cor_laranja}, 'valueformat': '.3f'},
        gauge={'axis': {'range': [0, 1.0], 'tickfont': {'size': 14}}, 
               'bar': {'color': cor_laranja}, 
               'threshold': {'line': {'color': "black", 'width': 5}, 'thickness': 0.8, 'value': meta_valor}}
    ), row=2, col=1)

    # --- 3. OFENSORES ---
    ofensores = pd.Series({
        'R. Grandes': linha_atual['rogr_dev'], 
        'R. Médias': linha_atual['rome_dev'], 
        'R. Pequenas': linha_atual['rope_dev']
    }).sort_values()

    fig.add_trace(go.Bar(
        x=ofensores.values, y=ofensores.index, orientation='h',
        marker_color=cor_laranja,
        text=ofensores.values.astype(int), textposition='outside',
        textfont=dict(size=20, family="Arial Black")
    ), row=2, col=2)

    # --- 4. MIX PRODUÇÃO (AJUSTADO PARA NÃO CHOCAR) ---
    labels_mix = []
    valores_mix = []
    for col in colunas_producao:
        if linha_atual[col] > 0:
            # Limpeza e abreviação para nomes muito longos
            nome = col.replace('_ent', '').replace('ro_', '').upper()
            nome = nome.replace('TIRAS DE ANEL', 'TIRAS') 
            labels_mix.append(nome)
            valores_mix.append(int(linha_atual[col]))

    fig.add_trace(go.Bar(
        x=labels_mix, y=valores_mix,
        marker_color=CONFIG["Cores"]["Primaria"],
        text=valores_mix, textposition='outside',
        textfont=dict(size=16, family="Arial Black")
    ), row=3, col=1)

    # --- 5. TENDÊNCIA HISTÓRICA ---
    fig.add_trace(go.Scatter(
        x=df['mes_curto'], y=df['perc_dev'],
        mode='lines+markers+text',
        line=dict(color=CONFIG["Cores"]["Primaria"], width=5),
        marker=dict(size=12),
        text=[f"{v:.2f}%" for v in df['perc_dev']], textposition="top center",
        textfont=dict(size=14, family="Arial Black", color=CONFIG["Cores"]["Primaria"]),
        name="Realizado"
    ), row=3, col=2)

    fig.add_trace(go.Scatter(
        x=df['mes_curto'], y=[meta_valor] * len(df),
        mode='lines+text',
        line=dict(color=cor_laranja, dash='dash', width=3),
        text=[f"META {meta_valor:.2f}%"] + [""] * (len(df)-1),
        textposition="top right",
        textfont=dict(size=15, family="Arial Black", color=cor_laranja),
        name="Meta"
    ), row=3, col=2)

    # --- AJUSTES FINAIS ---
    fig.update_layout(
        title=f"RELATÓRIO ÍNDICE DE DEVOLUÇÕES - {mes_nome} / 2026",
        title_font=dict(size=36, family="Arial Black", color=CONFIG["Cores"]["Primaria"]),
        title_x=0.5,
        width=CONFIG["Largura_Pixels"], height=CONFIG["Altura_Pixels"],
        template="plotly_white",
        margin=dict(t=150, b=120, l=80, r=50), # b=120 dá espaço para a inclinação
        showlegend=False
    )

    # Ajuste específico do Eixo X do Mix de Produção
    fig.update_xaxes(
        tickangle=-20, 
        tickfont=dict(size=12, family="Arial Black"), 
        row=3, col=1
    )
    
    # Ajuste do Eixo X da Evolução (mantém reto)
    fig.update_xaxes(tickangle=0, tickfont=dict(size=14, family="Arial Black"), row=3, col=2)

    for annotation in fig['layout']['annotations']:
        annotation['font'] = dict(size=20, color="#333", family="Arial Black")

    if not os.path.exists(CONFIG["Diretorio_Saida"]): os.makedirs(CONFIG["Diretorio_Saida"])
    caminho_final = f"{CONFIG['Diretorio_Saida']}/Dashboard_Qualidade_A3_{mes_nome}.png"
    
    fig.write_image(caminho_final, scale=CONFIG["DPI_Scale"])
    print(f"\n[OK] Dashboard gerado com correção de eixos: {caminho_final}")

if __name__ == "__main__":
    dados_df, colunas_filtradas = carregar_e_processar()
    if dados_df is not None:
        gerar_relatorio_a3(dados_df, colunas_filtradas)
