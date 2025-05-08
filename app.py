import dash
from dash import html, dash_table
import pandas as pd

# Cargar el archivo Excel
df = pd.read_excel("ejemplo.xlsx")  # Reemplaza con el nombre correcto del archivo

app = dash.Dash(__name__)

app.layout = html.Div(children=[
    html.H1("Dashboard de Datos desde Excel"),
    dash_table.DataTable(
        data=df.to_dict('records'),
        columns=[{"name": i, "id": i} for i in df.columns],
        page_size=35,
        style_table={'overflowX': 'auto'},
        style_cell={'textAlign': 'left'}
    )
])

# Paso 1: Cargar los distintos bloques de tablas del Excel
# Ajustá 'skiprows' y 'nrows' a los bloques reales de tu archivo
tabla_squall = pd.read_excel("ejemplo.xlsx", skiprows=2, nrows=12).iloc[:, 7:]
tabla_venc_1_4 = pd.read_excel("ejemplo.xlsx", skiprows=17, nrows=10).iloc[:, 1:]
tabla_venc_5_6 = pd.read_excel("ejemplo.xlsx", skiprows=30, nrows=8)
tabla_venc_7_8 = pd.read_excel("ejemplo.xlsx", skiprows=40, nrows=8)

# Paso 2: Crear la app y el layout
app = dash.Dash(__name__)

app.layout = html.Div([
    html.H2("SQUALL ANA", style={'textAlign': 'center'}),
    dash_table.DataTable(
        data=tabla_squall.to_dict('records'),
        columns=[{"name": i, "id": i} for i in tabla_squall.columns],
        page_size=15
    ),

    html.H2("VENCIMIENTO ETAPAS 1-4", style={'textAlign': 'center', 'marginTop': '40px'}),
    dash_table.DataTable(
        data=tabla_venc_1_4.to_dict('records'),
        columns=[{"name": i, "id": i} for i in tabla_venc_1_4.columns],
        page_size=10
    ),

    html.H2("VENCIMIENTO ETAPAS 5-6", style={'textAlign': 'center', 'marginTop': '40px'}),
    dash_table.DataTable(
        data=tabla_venc_5_6.to_dict('records'),
        columns=[{"name": i, "id": i} for i in tabla_venc_5_6.columns],
        page_size=10
    ),

    html.H2("VENCIMIENTO ETAPAS 7-8", style={'textAlign': 'center', 'marginTop': '40px'}),
    dash_table.DataTable(
        data=tabla_venc_7_8.to_dict('records'),
        columns=[{"name": i, "id": i} for i in tabla_venc_7_8.columns],
        page_size=10
    )
])

if __name__ == "__main__":
    app.run(debug=True)