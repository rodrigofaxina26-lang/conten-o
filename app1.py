import dash
import pandas as pd
import dash
from dash import dcc, html, Input, Output
import plotly.graph_objects as go
import plotly.express as px

# --- CONFIGURAÇÕES ---
CAMINHO = r"P:\PUBLICO\Contenção\Quarentena Controle\2026\Lista contenção_ redução. revisao 06_01_2026.xlsx"
ABA_DADOS = "dados_dashboard"

app = dash.Dash(__name__, title="Dashboard Qualidade Pro-Metal", external_stylesheets=['https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&display=swap'])

def carregar_dados_auxiliares():
    try:
        # Re-lê o arquivo (o Python busca a versão salva no disco)
        df = pd.read_excel(CAMINHO, sheet_name=ABA_DADOS)
        df.columns = [str(c).strip() for c in df.columns]
        
        df['Inspeção'] = pd.to_numeric(df['Inspeção'], errors='coerce').fillna(0)
        df['NC'] = pd.to_numeric(df['NC'], errors='coerce').fillna(0)
        
        df['Codigo'] = df['Codigo'].astype(str).str.strip()
        df['Status Insp.'] = df['Status Insp.'].astype(str).str.strip().str.upper()
        
        # Ignora o que não tem inspeção ou status for '0'
        df = df[df['Inspeção'] > 0]
        df = df[~df['Status Insp.'].isin(['NAN', '0', '0.0', 'NONE', ''])]
        
        return df

    except Exception as e:
        print(f"Erro ao ler a aba: {e}")
        return pd.DataFrame()

# --- LAYOUT COM CORES CLARAS ---
app.layout = html.Div(className='container', children=[
    
    # Relógio de atualização automática (5 minutos = 300.000 ms)
    dcc.Interval(id='intervalo-atualizacao', interval=300000, n_intervals=0),

    # HEADER
    html.Div(className='header', children=[
        html.Img(src='/assets/logo.png', className='logo'),
        html.H1("📊 CONTROLE DE QUALIDADE PRO-METAL", className='title'),
        html.P("Atualização automática a cada 5 minutos", className='subtitle')
    ]),

    html.Div(children=[
        # Pizza
        html.Div(className='pie-chart chart-container', children=[
            dcc.Graph(id='grafico-pizza')
        ]),
        
        # Seletor
        html.Div(className='filter-container', children=[
            html.Label("Filtro de Status Ativos:", className='filter-label'),
            dcc.Dropdown(id='selecao-status', style={'marginTop': '10px'})
        ])
    ], style={'marginBottom': '30px'}),

    # Barras
    html.Div(className='bar-chart chart-container', children=[
        dcc.Graph(id='grafico-barras-principal')
    ])
])

# --- LÓGICA DE ATUALIZAÇÃO ---
@app.callback(
    [Output('grafico-pizza', 'figure'),
     Output('grafico-barras-principal', 'figure'),
     Output('selecao-status', 'options'),
     Output('selecao-status', 'value')],
    [Input('selecao-status', 'value'),
     Input('intervalo-atualizacao', 'n_intervals')] # Segunda entrada: o tempo passando
)
def atualizar_dash(status_filtro, n):

    df = carregar_dados_auxiliares()
    
    if df.empty:
        fig_vazia = go.Figure().add_annotation(text="Nenhum dado ativo no momento", showarrow=False)
        return fig_vazia, fig_vazia, [], None

    status_unicos = sorted([str(s) for s in df['Status Insp.'].unique()])
    opcoes_status = [{'label': s, 'value': s} for s in status_unicos]
    
    if not status_filtro or status_filtro not in status_unicos:
        status_filtro = status_unicos[0]

    # Gráfico de Pizza (Cores Claras)
    df_pizza = df.groupby('Status Insp.')['Inspeção'].sum().reset_index()
    fig_pizza = px.pie(df_pizza, values='Inspeção', names='Status Insp.', hole=0.4, template='plotly_white')
    fig_pizza.update_layout(title="Distribuição por Status", height=700)

    # Gráfico de Barras
    df_filtrado = df[df['Status Insp.'] == status_filtro].sort_values('Inspeção', ascending=False).head(25)
    
    fig_barras = go.Figure()

    fig_barras.add_trace(go.Bar(
        x=df_filtrado['Codigo'], y=df_filtrado['Inspeção'], 
        name="Inspecionado", marker_color='#3b82f6', # Azul profissional
        text=df_filtrado['Inspeção'], textposition='outside'
    ))

    fig_barras.add_trace(go.Scatter(
        x=df_filtrado['Codigo'], y=df_filtrado['NC'], 
        name="NC", marker_color='#ef4444', # Vermelho NC
        mode='lines+markers+text', text=df_filtrado['NC'], 
        textposition='top center', yaxis='y2'
    ))

    fig_barras.update_layout(
        template='plotly_white',
        xaxis={'type': 'category', 'title': 'Código do Produto'},
        yaxis={'title': 'Volume de Inspeção'},
        yaxis2={'title': 'Volume de NC', 'overlaying': 'y', 'side': 'right'},
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=80)
    )

    return fig_pizza, fig_barras, opcoes_status, status_filtro

# --- EXECUÇÃO ---
if __name__ == '__main__':
    app.run(debug=False, host='192.168.1.150', port=5005)