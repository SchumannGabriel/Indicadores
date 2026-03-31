# Gerador de Indicadores para Impressão (A3)

Este repositório contém scripts em Python desenvolvidos para automatizar a criação de indicadores de performance, otimizados especificamente para **impressão em formato A3**. O objetivo é facilitar a gestão visual em ambientes industriais ou administrativos através de dashboards estáticos de alta resolução.

##  Funcionalidades

O projeto processa dados (via `.xlsx` ou `.txt`) e gera visualizações profissionais para:
* **OTD (On-Time Delivery):** Relatórios de pontualidade de entrega.
* **Absenteísmo:** Controle de frequência e histórico.
* **Custo por Funcionário:** Análise financeira operacional.
* **Não Conformidades (NC):** Monitoramento de processos internos.
* **Índice de Devoluções:** Qualidade e logística.

## Configuração para Impressão A3

Para garantir que os gráficos fiquem nítidos em uma folha A3 ($297mm \times 420mm$), os scripts utilizam as seguintes premissas técnicas:

1.  **DPI Alto:** Configurado para no mínimo 300 DPI para evitar serrilhamento.
2.  **Dimensões da Figura:** Ajustadas via `matplotlib` para o aspect ratio do padrão ISO 216.
3.  **Margens:** Otimizadas para não haver corte de informações pelas impressoras térmicas ou laser comuns.

> **Dica:** Ao salvar seus gráficos, use:
> ```python
> plt.savefig('indicador_a3.png', dpi=300, bbox_inches='tight')
> ```

##  Tecnologias Utilizadas

* **Python 3.x**
* **Pandas:** Manipulação de dados.
* **Matplotlib / Seaborn:** Geração de gráficos e layouts.
* **Openpyxl:** Leitura de arquivos Excel.

## Estrutura do Repositório

* `/diarios`: Scripts para indicadores de atualização diária.
* `OTD_geral_impressao.py`: Script principal configurado para o layout de impressão.

##  Como Executar

1. Clone o repositório:
   ```bash
   git clone [https://github.com/SchumannGabriel/Indicadores.git](https://github.com/SchumannGabriel/Indicadores.git)
   ```
2. Instale as dependências:
```bash
pip install pandas matplotlib seaborn openpyxl
```
3. Certifique-se de que o arquivo Dados.xlsx está atualizado.
4. Execute o script desejado:
```bash
python OTD_geral_impressao.py
```

# Notas de Versão
### v1.0: Implementação dos indicadores base e ajuste de escala para A3.
## Desenvolvido por Gabriel Schumann
