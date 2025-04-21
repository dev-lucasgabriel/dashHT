
import pandas as pd
import dash
import dash_bootstrap_components as dbc
from dash import dcc, html, dash_table
from dash.dependencies import Input, Output
import plotly.graph_objects as go

arquivo = "ENTREGAS FOLHA 2.xlsx"

def carregar_dados():
    xls = pd.ExcelFile(arquivo)
    dados_lista = []

    for mes in xls.sheet_names:
        try:
            df = pd.read_excel(arquivo, sheet_name=mes, usecols="C:G", skiprows=3, nrows=91)
            df.columns = ["Empresa", "Recibo", "Remuneracao", "DCTF", "INSS"]
            df = df[df["Empresa"].notna()]
            df["Mês"] = mes
            dados_lista.append(df)
        except Exception as e:
            print(f"Aba {mes} ignorada: {e}")
            continue

    if not dados_lista:
        print("⚠️ Nenhum dado válido encontrado.")
        return pd.DataFrame(columns=["Empresa", "Recibo", "Remuneracao", "DCTF", "INSS", "Mês"])

    return pd.concat(dados_lista, ignore_index=True)

dados = carregar_dados()

# Cálculo da conformidade por empresa
status_cols = ["Recibo", "Remuneracao", "DCTF", "INSS"]

def calcular_conformidade(df):
    def porcentagem_ok(linha):
        total = len(status_cols)
        ok = sum([1 for col in status_cols if str(linha[col]).strip().upper() == "OK"])
        return round((ok / total) * 100, 2)
    df["%OK"] = df.apply(porcentagem_ok, axis=1)
    return df

dados = calcular_conformidade(dados)

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.FLATLY])

app.layout = dbc.Container([
    html.H2("Dashboard de Conformidade das Empresas", style={"textAlign": "center", "marginTop": "20px"}),

    dbc.Row([
        dbc.Col([
            html.Label("Selecione o Mês:"),
            dcc.Dropdown(
                id="filtro-mes",
                options=[{"label": m, "value": m} for m in sorted(dados["Mês"].unique())],
                value=sorted(dados["Mês"].unique())[0] if not dados.empty else None,
                clearable=False
            )
        ], width=4),

        dbc.Col([
            html.Label("Selecione a Empresa:"),
            dcc.Dropdown(
                id="filtro-empresa",
                options=[],
                multi=False
            )
        ], width=8)
    ], className="mb-4"),

    dbc.Row([
        dbc.Col([
            dash_table.DataTable(
                id='tabela-status',
                columns=[{"name": col, "id": col} for col in ["Empresa", "Recibo", "Remuneracao", "DCTF", "INSS", "%OK"]],
                style_cell={'textAlign': 'left', 'padding': '5px', 'color': 'black', 'backgroundColor': 'white'},
                style_header={'backgroundColor': 'white', 'fontWeight': 'bold'},
                style_table={'overflowX': 'auto'},
                page_size=25
            )
        ])
    ], className="mb-4"),

    dbc.Row([
        dbc.Col([
            html.H5("Empresas com Pendências"),
            dash_table.DataTable(
                id='empresas-incompletas',
                columns=[
                    {"name": "Empresa", "id": "Empresa"},
                    {"name": "%OK", "id": "%OK"}
                ],
                data=[],
                style_cell={'textAlign': 'left', 'color': 'black', 'backgroundColor': 'white'},
                style_header={'backgroundColor': 'white', 'fontWeight': 'bold'},
                style_table={'overflowX': 'auto'},
                page_size=10
            )
        ])
    ])
])

@app.callback(
    Output("filtro-empresa", "options"),
    Output("filtro-empresa", "value"),
    Input("filtro-mes", "value")
)
def atualiza_empresas(mes):
    empresas = dados[dados["Mês"] == mes]["Empresa"].unique()
    opcoes = [{"label": e, "value": e} for e in sorted(empresas)]
    return opcoes, None

@app.callback(
    Output("tabela-status", "data"),
    Output("empresas-incompletas", "data"),
    Input("filtro-mes", "value"),
    Input("filtro-empresa", "value")
)
def atualizar_tabelas(mes, empresa):
    df_mes = dados[dados["Mês"] == mes]
    if empresa:
        df_emp = df_mes[df_mes["Empresa"] == empresa]
    else:
        df_emp = pd.DataFrame()

    empresas_incompletas = df_mes[df_mes["%OK"] < 100][["Empresa", "%OK"]].drop_duplicates()

    return df_emp.to_dict("records"), empresas_incompletas.to_dict("records")

if __name__ == '__main__':
    app.run(debug=True, port=8050, host='0.0.0.0')
