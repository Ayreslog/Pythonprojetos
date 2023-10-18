import pandas as pd
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import plotly.express as px

# Leitura dos dados
df = pd.read_excel("C:\\Users\\ayres.filho\\Desktop\\Auditoria\\resultado_divergencias_consolidados.xlsx")

# Correção dos valores não numéricos
df['Divergencias'] = pd.to_numeric(df['Divergencias'], errors='coerce')

# Filtrar itens com divergência e dados na coluna data_ajuste
df_com_divergencia = df[(df['Divergencias'] != 0) & df['data_ajuste'].notna()]

# Inicialização do aplicativo Dash
app = dash.Dash(__name__)

# Layout do dashboard
app.layout = html.Div([
    html.H1("Dashboard de Divergências", style={'text-align': 'center', 'text-transform': 'uppercase'}),

    # Filtros
    html.Div([
        dcc.Dropdown(
            id='filtro-auditor',
            options=[{'label': auditor, 'value': auditor} for auditor in df['AUDITOR'].unique()],
            multi=True,
            value=[]
        ),
        dcc.Dropdown(
            id='filtro-ga',
            options=[{'label': ga, 'value': ga} for ga in df['GA'].unique()],
            multi=True,
            value=[]
        ),
    ]),

    # Gráfico das semanas com mais divergências
    dcc.Graph(id='grafico-semanas', config={'displayModeBar': False}),

    # Gráfico dos dias com mais divergências por semana
    dcc.Graph(id='grafico-dias-semana', config={'displayModeBar': False}),

    # Gráfico de divergências por unidades
    dcc.Graph(id='grafico-unidades', config={'displayModeBar': False}),

    # Gráfico de divergências por itens
    dcc.Graph(id='grafico-itens', config={'displayModeBar': False}),
])

# Função para criar gráfico de barras
def criar_grafico_barras(data, x, y, title):
    fig = px.bar(data, x=x, y=y, title=title, text=y)
    return fig

# Callback para atualizar os gráficos com base nos filtros e na seleção interativa
@app.callback(
    [Output('grafico-semanas', 'figure'),
     Output('grafico-dias-semana', 'figure'),
     Output('grafico-unidades', 'figure'),
     Output('grafico-itens', 'figure')],
    [Input('filtro-auditor', 'value'),
     Input('filtro-ga', 'value'),
     Input('grafico-semanas', 'relayoutData')]  # Adicionamos a entrada para a seleção interativa
)
def update_graficos(filtros_auditor, filtros_ga, relayoutData):
    df_filtrado = df_com_divergencia

    if filtros_auditor:
        df_filtrado = df_filtrado[df_filtrado['AUDITOR'].isin(filtros_auditor)]

    if filtros_ga:
        df_filtrado = df_filtrado[df_filtrado['GA'].isin(filtros_ga)]

    # Verifica se houve uma seleção no gráfico de semanas
    if relayoutData and 'xaxis.range[0]' in relayoutData:
        x_min = relayoutData['xaxis.range[0]']
        x_max = relayoutData['xaxis.range[1]']
        # Filtra o DataFrame com base na seleção no gráfico de semanas
        df_filtrado = df_filtrado[(df_filtrado['SEMANAS'] >= x_min) & (df_filtrado['SEMANAS'] <= x_max)]

    # Contagem de divergências por semana
    contagem_semanas = df_filtrado['SEMANAS'].value_counts().reset_index()
    contagem_semanas.columns = ['Semana', 'Contagem']

    # Contagem de divergências por dia da semana
    contagem_dias_semana = df_filtrado['DIA DA SEMANA'].value_counts().reset_index()
    contagem_dias_semana.columns = ['Dia da Semana', 'Contagem']

    # Contagem de divergências por unidades
    contagem_unidades = df_filtrado['Unidades 33/Ciclica Padrão'].value_counts().reset_index()
    contagem_unidades.columns = ['Unidades', 'Contagem']

    # Contagem de divergências por itens
    contagem_itens = df_filtrado['ITEM'].value_counts().reset_index().head(5)
    contagem_itens.columns = ['Item', 'Contagem']

    return criar_grafico_barras(contagem_semanas, 'Semana', 'Contagem', 'Semanas com Mais Divergências'), \
           criar_grafico_barras(contagem_dias_semana, 'Dia da Semana', 'Contagem',
                                'Dias com Mais Divergências na Semana'), \
           criar_grafico_barras(contagem_unidades, 'Unidades', 'Contagem', 'Divergências por Unidades'), \
           criar_grafico_barras(contagem_itens, 'Item', 'Contagem', 'Itens com Mais Divergências')

if __name__ == '__main__':
    app.run_server(debug=True)
