"""Dashboard interactivo con Plotly Dash para Search Console Ruta del archivo

Excel: C:\Dash\Search-Console\search-consoletl.xlsx
"""

from datetime import datetime
import os
import dash
from dash import Input, Output, dash_table, dcc, html
import dash_bootstrap_components as dbc
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# Rutas de archivo
DEFAULT_PATH = r"C:\Dash\Search-Console\search-consoletl.xlsx"
FALLBACK_PATH = "search-consoletl.xlsx"

file_path = DEFAULT_PATH if os.path.exists(DEFAULT_PATH) else FALLBACK_PATH
if not os.path.exists(file_path):
  df_grafico = pd.DataFrame(
      columns=["Fecha", "Clics", "Impresiones", "CTR", "Posición"]
  )
  df_consultas = pd.DataFrame(
      columns=["Consultas principales", "Clics", "Impresiones", "CTR", "Posición"]
  )
  df_paginas = pd.DataFrame(
      columns=["Páginas principales", "Clics", "Impresiones", "CTR", "Posición"]
  )
  df_paises = pd.DataFrame(
      columns=["País", "Clics", "Impresiones", "CTR", "Posición"]
  )
  df_dispositivos = pd.DataFrame(
      columns=["Dispositivo", "Clics", "Impresiones", "CTR", "Posición"]
  )
  df_filtros = pd.DataFrame(columns=["Filtrar", "Valor"])
  min_date = datetime(2026, 1, 1).date()
  max_date = datetime(2026, 12, 31).date()
else:
  xls = pd.ExcelFile(file_path)
  df_grafico = (
      pd.read_excel(xls, sheet_name="Gráfico")
      if "Gráfico" in xls.sheet_names
      else pd.DataFrame()
  )
  df_consultas = (
      pd.read_excel(xls, sheet_name="Consultas")
      if "Consultas" in xls.sheet_names
      else pd.DataFrame()
  )
  df_paginas = (
      pd.read_excel(xls, sheet_name="Páginas")
      if "Páginas" in xls.sheet_names
      else pd.DataFrame()
  )
  df_paises = (
      pd.read_excel(xls, sheet_name="Países")
      if "Países" in xls.sheet_names
      else pd.DataFrame()
  )
  df_dispositivos = (
      pd.read_excel(xls, sheet_name="Dispositivos")
      if "Dispositivos" in xls.sheet_names
      else pd.DataFrame()
  )
  df_filtros = (
      pd.read_excel(xls, sheet_name="Filtros")
      if "Filtros" in xls.sheet_names
      else pd.DataFrame()
  )

  if not df_grafico.empty and "Fecha" in df_grafico.columns:
    df_grafico["Fecha"] = pd.to_datetime(df_grafico["Fecha"])
    df_grafico = df_grafico.sort_values("Fecha").reset_index(drop=True)
    min_date = df_grafico["Fecha"].min().date()
    max_date = df_grafico["Fecha"].max().date()
  else:
    min_date = datetime(2026, 1, 1).date()
    max_date = datetime(2026, 12, 31).date()

# Inicialización de Dash con Bootstrap
app = dash.Dash(
    __name__,
    external_stylesheets=[dbc.themes.FLATLY, dbc.icons.BOOTSTRAP],
    title="Google Search Console Analytics | Dash",
)

CARD_STYLE = {
    "boxShadow": "0 4px 6px -1px rgba(0,0,0,0.08)",
    "borderRadius": "12px",
    "border": "1px solid #E2E8F0",
    "backgroundColor": "#FFFFFF",
    "padding": "1.2rem",
    "height": "100%",
}

# Layout de la aplicación
app.layout = dbc.Container(
    [
        # Encabezado
        dbc.Row(
            [
                dbc.Col(
                    [
                        html.Div(
                            [
                                html.H2(
                                    [
                                        html.I(
                                            className=(
                                                "bi bi-google me-2 text-primary"
                                            )
                                        ),
                                        "Google Search Console Analytics",
                                    ],
                                    className="fw-bold mb-1 text-dark",
                                ),
                                html.P(
                                    "Dashboard interactivo para Search Console"
                                    " (tradinglatinoacademy.com) | Archivo:"
                                    f" {os.path.basename(file_path)}",
                                    className="text-muted mb-0",
                                ),
                            ],
                            className="py-3",
                        )
                    ],
                    width=12,
                )
            ],
            className="border-bottom mb-4",
        ),
        # Panel de Controles: Selector de Fechas y Sliders
        dbc.Card(
            [
                dbc.CardBody([
                    dbc.Row(
                        [
                            dbc.Col(
                                [
                                    html.Label(
                                        "🗓️ Rango de Fechas:",
                                        className=(
                                            "fw-bold text-secondary mb-1"
                                        ),
                                    ),
                                    dcc.DatePickerRange(
                                        id="date-picker-range",
                                        min_date_allowed=min_date,
                                        max_date_allowed=max_date,
                                        start_date=min_date,
                                        end_date=max_date,
                                        display_format="YYYY-MM-DD",
                                        style={"width": "100%"},
                                    ),
                                ],
                                xs=12,
                                md=6,
                                lg=4,
                            ),
                            dbc.Col(
                                [
                                    html.Label(
                                        "🔢 Top N barras a mostrar:",
                                        className=(
                                            "fw-bold text-secondary mb-1"
                                        ),
                                    ),
                                    dcc.Slider(
                                        id="top-n-slider",
                                        min=5,
                                        max=30,
                                        step=5,
                                        value=10,
                                        marks={
                                            5: "5",
                                            10: "10",
                                            15: "15",
                                            20: "20",
                                            25: "25",
                                            30: "30",
                                        },
                                        tooltip={
                                            "placement": "bottom",
                                            "always_visible": False,
                                        },
                                    ),
                                ],
                                xs=12,
                                md=6,
                                lg=4,
                            ),
                            dbc.Col(
                                [
                                    html.Label(
                                        "ℹ️ Filtros de Búsqueda:",
                                        className=(
                                            "fw-bold text-secondary mb-1"
                                        ),
                                    ),
                                    html.Div([
                                        dbc.Badge(
                                            f"{r['Filtrar']}: {r['Valor']}",
                                            color="info",
                                            className="me-2 mb-1 p-2",
                                        )
                                        for _, r in df_filtros.iterrows()
                                    ] if not df_filtros.empty else html.Span(
                                        "Sin metadatos", className="text-muted"
                                    )),
                                ],
                                xs=12,
                                md=12,
                                lg=4,
                            ),
                        ],
                        className="align-items-center",
                    )
                ])
            ],
            className="mb-4 shadow-sm border-0 bg-light",
        ),
        # Fila de Tarjetas de KPIs reactivos
        html.Div(id="kpi-container", className="mb-4"),
        # Pestañas / Tabs
        dbc.Tabs(
            [
                # Tab 1: Evolución Diaria
                dbc.Tab(
                    [
                        dbc.Row(
                            [
                                dbc.Col(
                                    [
                                        dbc.Card(
                                            [
                                                dbc.CardHeader(
                                                    [
                                                        html.Div(
                                                            [
                                                                html.Span(
                                                                    "Evolución"
                                                                    " Diaria de"
                                                                    " Métricas",
                                                                    className=(
                                                                        "fw-bold"
                                                                    ),
                                                                ),
                                                                dcc.RadioItems(
                                                                    id=(
                                                                        "metric-radio"
                                                                    ),
                                                                    options=[
                                                                        {
                                                                            "label": (
                                                                                " Clics"
                                                                                " e"
                                                                                " Impresiones"
                                                                                " (Doble"
                                                                                " Eje)"
                                                                            ),
                                                                            "value": (
                                                                                "dual"
                                                                            ),
                                                                        },
                                                                        {
                                                                            "label": (
                                                                                " Solo"
                                                                                " Clics"
                                                                            ),
                                                                            "value": (
                                                                                "clics"
                                                                            ),
                                                                        },
                                                                        {
                                                                            "label": (
                                                                                " Solo"
                                                                                " Impresiones"
                                                                            ),
                                                                            "value": (
                                                                                "impresiones"
                                                                            ),
                                                                        },
                                                                        {
                                                                            "label": (
                                                                                " CTR"
                                                                                " (%)"
                                                                            ),
                                                                            "value": (
                                                                                "ctr"
                                                                            ),
                                                                        },
                                                                        {
                                                                            "label": (
                                                                                " Posición"
                                                                                " Media"
                                                                            ),
                                                                            "value": (
                                                                                "posicion"
                                                                            ),
                                                                        },
                                                                    ],
                                                                    value="dual",
                                                                    inline=True,
                                                                    inputClassName=(
                                                                        "me-1"
                                                                        " ms-3"
                                                                    ),
                                                                    labelClassName=(
                                                                        "text-secondary"
                                                                        " small"
                                                                        " fw-semibold"
                                                                    ),
                                                                ),
                                                            ],
                                                            className=(
                                                                "d-flex"
                                                                " flex-wrap"
                                                                " justify-content-between"
                                                                " align-items-center"
                                                            ),
                                                        )
                                                    ],
                                                    className=(
                                                        "bg-white border-bottom"
                                                    ),
                                                ),
                                                dbc.CardBody([
                                                    dcc.Graph(
                                                        id="daily-time-series",
                                                        config={
                                                            "displayModeBar": (
                                                                True
                                                            ),
                                                            "responsive": True,
                                                        },
                                                    )
                                                ]),
                                            ],
                                            style=CARD_STYLE,
                                        )
                                    ],
                                    width=12,
                                )
                            ],
                            className="mt-3",
                        )
                    ],
                    label="📊 Rendimiento Diario",
                    tab_id="tab-rendimiento",
                ),
                # Tab 2: Comparativas Semanal y Mensual (Barras)
                dbc.Tab(
                    [
                        dbc.Row(
                            [
                                dbc.Col(
                                    [
                                        dbc.Card(
                                            [
                                                dbc.CardHeader(
                                                    "📆 Comparativa Mensual"
                                                    " (MoM)",
                                                    className=(
                                                        "fw-bold bg-white"
                                                    ),
                                                ),
                                                dbc.CardBody([
                                                    dcc.Graph(
                                                        id="chart-monthly-clics"
                                                    ),
                                                    dcc.Graph(
                                                        id="chart-monthly-imp"
                                                    ),
                                                    html.H6(
                                                        "Tabla Resumen Mensual",
                                                        className=(
                                                            "fw-bold mt-3"
                                                            " text-secondary"
                                                        ),
                                                    ),
                                                    html.Div(
                                                        id=(
                                                            "table-monthly-container"
                                                        )
                                                    ),
                                                ]),
                                            ],
                                            style=CARD_STYLE,
                                        )
                                    ],
                                    xs=12,
                                    lg=6,
                                    className="mb-3",
                                ),
                                dbc.Col(
                                    [
                                        dbc.Card(
                                            [
                                                dbc.CardHeader(
                                                    "📅 Comparativa Semanal"
                                                    " (WoW)",
                                                    className=(
                                                        "fw-bold bg-white"
                                                    ),
                                                ),
                                                dbc.CardBody([
                                                    dcc.Graph(
                                                        id="chart-weekly-clics"
                                                    ),
                                                    dcc.Graph(
                                                        id="chart-weekly-imp"
                                                    ),
                                                    html.H6(
                                                        "Tabla Resumen Semanal",
                                                        className=(
                                                            "fw-bold mt-3"
                                                            " text-secondary"
                                                        ),
                                                    ),
                                                    html.Div(
                                                        id=(
                                                            "table-weekly-container"
                                                        )
                                                    ),
                                                ]),
                                            ],
                                            style=CARD_STYLE,
                                        )
                                    ],
                                    xs=12,
                                    lg=6,
                                    className="mb-3",
                                ),
                            ],
                            className="mt-3",
                        )
                    ],
                    label="📅 Comparativas Semanal / Mensual",
                    tab_id="tab-comparativas",
                ),
                # Tab 3: Consultas (Keywords)
                dbc.Tab(
                    [
                        dbc.Row(
                            [
                                dbc.Col(
                                    [
                                        dbc.Card(
                                            [
                                                dbc.CardBody([
                                                    dbc.Input(
                                                        id="input-kw-search",
                                                        type="text",
                                                        placeholder=(
                                                            "🔎 Filtrar"
                                                            " consultas por"
                                                            " palabra clave..."
                                                        ),
                                                        className="mb-3",
                                                    ),
                                                    dbc.Row([
                                                        dbc.Col(
                                                            [
                                                                dcc.Graph(
                                                                    id=(
                                                                        "chart-kw-clics"
                                                                    )
                                                                )
                                                            ],
                                                            xs=12,
                                                            md=6,
                                                        ),
                                                        dbc.Col(
                                                            [
                                                                dcc.Graph(
                                                                    id=(
                                                                        "chart-kw-imp"
                                                                    )
                                                                )
                                                            ],
                                                            xs=12,
                                                            md=6,
                                                        ),
                                                    ]),
                                                    html.H6(
                                                        "Listado Completo de"
                                                        " Consultas",
                                                        className=(
                                                            "fw-bold mt-4 mb-2"
                                                            " text-secondary"
                                                        ),
                                                    ),
                                                    html.Div(
                                                        id="table-kw-container"
                                                    ),
                                                ])
                                            ],
                                            style=CARD_STYLE,
                                        )
                                    ],
                                    width=12,
                                )
                            ],
                            className="mt-3",
                        )
                    ],
                    label="🔍 Consultas (Keywords)",
                    tab_id="tab-consultas",
                ),
                # Tab 4: Páginas Principales
                dbc.Tab(
                    [
                        dbc.Row(
                            [
                                dbc.Col(
                                    [
                                        dbc.Card(
                                            [
                                                dbc.CardBody([
                                                    dbc.Row([
                                                        dbc.Col(
                                                            [
                                                                dcc.Graph(
                                                                    id=(
                                                                        "chart-pages-clics"
                                                                    )
                                                                )
                                                            ],
                                                            xs=12,
                                                            md=6,
                                                        ),
                                                        dbc.Col(
                                                            [
                                                                dcc.Graph(
                                                                    id=(
                                                                        "chart-pages-imp"
                                                                    )
                                                                )
                                                            ],
                                                            xs=12,
                                                            md=6,
                                                        ),
                                                    ]),
                                                    html.H6(
                                                        "Detalle de Páginas y"
                                                        " Métricas",
                                                        className=(
                                                            "fw-bold mt-4 mb-2"
                                                            " text-secondary"
                                                        ),
                                                    ),
                                                    html.Div(
                                                        id=(
                                                            "table-pages-container"
                                                        )
                                                    ),
                                                ])
                                            ],
                                            style=CARD_STYLE,
                                        )
                                    ],
                                    width=12,
                                )
                            ],
                            className="mt-3",
                        )
                    ],
                    label="📄 Páginas Principales",
                    tab_id="tab-paginas",
                ),
                # Tab 5: Países y Dispositivos
                dbc.Tab(
                    [
                        dbc.Row(
                            [
                                dbc.Col(
                                    [
                                        dbc.Card(
                                            [
                                                dbc.CardHeader(
                                                    "🌍 Tráfico por País",
                                                    className=(
                                                        "fw-bold bg-white"
                                                    ),
                                                ),
                                                dbc.CardBody([
                                                    dcc.Graph(
                                                        id="chart-countries"
                                                    )
                                                ]),
                                            ],
                                            style=CARD_STYLE,
                                        )
                                    ],
                                    xs=12,
                                    lg=7,
                                    className="mb-3",
                                ),
                                dbc.Col(
                                    [
                                        dbc.Card(
                                            [
                                                dbc.CardHeader(
                                                    "📱 Clics por Dispositivo",
                                                    className=(
                                                        "fw-bold bg-white"
                                                    ),
                                                ),
                                                dbc.CardBody([
                                                    dcc.Graph(
                                                        id="chart-devices"
                                                    ),
                                                    html.H6(
                                                        "Resumen por"
                                                        " Dispositivo",
                                                        className=(
                                                            "fw-bold mt-3"
                                                            " text-secondary"
                                                        ),
                                                    ),
                                                    html.Div(
                                                        id=(
                                                            "table-devices-container"
                                                        )
                                                    ),
                                                ]),
                                            ],
                                            style=CARD_STYLE,
                                        )
                                    ],
                                    xs=12,
                                    lg=5,
                                    className="mb-3",
                                ),
                            ],
                            className="mt-3",
                        )
                    ],
                    label="🌍 Países & Dispositivos",
                    tab_id="tab-geo-dev",
                ),
            ],
            active_tab="tab-rendimiento",
        ),
    ],
    fluid=True,
    className="p-4 bg-light min-vh-100",
)


# =====================================================================
# CALLBACK 1: FILTRADO DE FECHAS, KPIS Y GRÁFICO DIARIO
# =====================================================================
@app.callback(
    [
        Output("kpi-container", "children"),
        Output("daily-time-series", "figure"),
    ],
    [
        Input("date-picker-range", "start_date"),
        Input("date-picker-range", "end_date"),
        Input("metric-radio", "value"),
    ],
)
def update_kpis_and_daily(start_date, end_date, selected_metric):
  if df_grafico.empty:
    return html.Div("No hay datos disponibles."), go.Figure()

  start_dt = pd.to_datetime(start_date)
  end_dt = pd.to_datetime(end_date)
  mask = (df_grafico["Fecha"] >= start_dt) & (df_grafico["Fecha"] <= end_dt)
  filtered = df_grafico[mask].copy()

  if filtered.empty:
    filtered = df_grafico.copy()

  # Variaciones (período actual vs período anterior de idéntica duración)
  n_days = len(filtered)
  if n_days >= 4:
    half = n_days // 2
    act = filtered.iloc[half:]
    ant = filtered.iloc[:half]

    c_act, c_ant = act["Clics"].sum(), ant["Clics"].sum()
    i_act, i_ant = act["Impresiones"].sum(), ant["Impresiones"].sum()

    delta_c = ((c_act - c_ant) / c_ant * 100) if c_ant > 0 else 0
    delta_i = ((i_act - i_ant) / i_ant * 100) if i_ant > 0 else 0

    ctr_act = (c_act / i_act * 100) if i_act > 0 else 0
    ctr_ant = (c_ant / i_ant * 100) if i_ant > 0 else 0
    delta_ctr = ctr_act - ctr_ant

    pos_act = act["Posición"].mean()
    pos_ant = ant["Posición"].mean()
    delta_pos = -(pos_act - pos_ant)
  else:
    delta_c = delta_i = delta_ctr = delta_pos = None

  tot_clics = filtered["Clics"].sum()
  tot_imp = filtered["Impresiones"].sum()
  tot_ctr = (tot_clics / tot_imp * 100) if tot_imp > 0 else 0
  mean_pos = filtered["Posición"].mean()

  def format_delta(d, unit="%"):
    if d is None:
      return ""
    sign = "+" if d > 0 else ""
    color = "text-success" if d >= 0 else "text-danger"
    return html.Small(
        f"{sign}{d:.1f}{unit} vs período anterior",
        className=f"d-block {color} fw-bold",
    )

  kpis_html = dbc.Row([
      dbc.Col(
          [
              dbc.Card(
                  [
                      dbc.CardBody([
                          html.P(
                              "🎯 Clics Totales",
                              className=(
                                  "text-muted text-uppercase small fw-bold mb-1"
                              ),
                          ),
                          html.H3(
                              f"{tot_clics:,}".replace(",", "."),
                              className="fw-bold mb-0 text-dark",
                          ),
                          format_delta(delta_c, "%"),
                      ])
                  ],
                  style=CARD_STYLE,
              )
          ],
          xs=12,
          sm=6,
          md=3,
          className="mb-2",
      ),
      dbc.Col(
          [
              dbc.Card(
                  [
                      dbc.CardBody([
                          html.P(
                              "👁️ Impresiones Totales",
                              className=(
                                  "text-muted text-uppercase small fw-bold mb-1"
                              ),
                          ),
                          html.H3(
                              f"{tot_imp:,}".replace(",", "."),
                              className="fw-bold mb-0 text-dark",
                          ),
                          format_delta(delta_i, "%"),
                      ])
                  ],
                  style=CARD_STYLE,
              )
          ],
          xs=12,
          sm=6,
          md=3,
          className="mb-2",
      ),
      dbc.Col(
          [
              dbc.Card(
                  [
                      dbc.CardBody([
                          html.P(
                              "📈 CTR Medio Ponderado",
                              className=(
                                  "text-muted text-uppercase small fw-bold mb-1"
                              ),
                          ),
                          html.H3(
                              f"{tot_ctr:.2f}%",
                              className="fw-bold mb-0 text-dark",
                          ),
                          format_delta(delta_ctr, " pts"),
                      ])
                  ],
                  style=CARD_STYLE,
              )
          ],
          xs=12,
          sm=6,
          md=3,
          className="mb-2",
      ),
      dbc.Col(
          [
              dbc.Card(
                  [
                      dbc.CardBody([
                          html.P(
                              "📍 Posición Media",
                              className=(
                                  "text-muted text-uppercase small fw-bold mb-1"
                              ),
                          ),
                          html.H3(
                              f"{mean_pos:.1f}",
                              className="fw-bold mb-0 text-dark",
                          ),
                          format_delta(delta_pos, " pts"),
                      ])
                  ],
                  style=CARD_STYLE,
              )
          ],
          xs=12,
          sm=6,
          md=3,
          className="mb-2",
      ),
  ])

  # Gráfica según la opción seleccionada
  if selected_metric == "dual":
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(
        go.Bar(
            x=filtered["Fecha"],
            y=filtered["Clics"],
            name="Clics",
            marker_color="#1E3A8A",
            opacity=0.85,
        ),
        secondary_y=False,
    )
    fig.add_trace(
        go.Scatter(
            x=filtered["Fecha"],
            y=filtered["Impresiones"],
            name="Impresiones",
            mode="lines+markers",
            line=dict(color="#06B6D4", width=2.5),
            marker=dict(size=4),
        ),
        secondary_y=True,
    )
    fig.update_layout(
        title="Evolución de Clics (Barras) e Impresiones (Línea)",
        hovermode="x unified",
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1
        ),
    )
    fig.update_yaxes(title_text="Clics", secondary_y=False)
    fig.update_yaxes(title_text="Impresiones", secondary_y=True)
  elif selected_metric == "clics":
    fig = px.bar(
        filtered,
        x="Fecha",
        y="Clics",
        title="Clics Diarios",
        color_discrete_sequence=["#1E3A8A"],
        text_auto=True,
    )
  elif selected_metric == "impresiones":
    fig = px.bar(
        filtered,
        x="Fecha",
        y="Impresiones",
        title="Impresiones Diarias",
        color_discrete_sequence=["#06B6D4"],
    )
  elif selected_metric == "ctr":
    filtered["CTR_pct"] = filtered["CTR"] * 100
    fig = px.line(
        filtered,
        x="Fecha",
        y="CTR_pct",
        title="CTR Diario (%)",
        markers=True,
        color_discrete_sequence=["#10B981"],
    )
    fig.update_yaxes(title_text="CTR (%)")
  else:
    fig = px.line(
        filtered,
        x="Fecha",
        y="Posición",
        title="Posición Media Diaria (Menor es mejor)",
        markers=True,
        color_discrete_sequence=["#F59E0B"],
    )
    fig.update_yaxes(autorange="reversed", title_text="Posición en Google")

  fig.update_layout(
      template="plotly_white", margin=dict(l=20, r=20, t=50, b=20), height=400
  )
  return kpis_html, fig


# =====================================================================
# CALLBACK 2: COMPARATIVAS SEMANALES Y MENSUALES (BARRAS)
# =====================================================================
@app.callback(
    [
        Output("chart-monthly-clics", "figure"),
        Output("chart-monthly-imp", "figure"),
        Output("table-monthly-container", "children"),
        Output("chart-weekly-clics", "figure"),
        Output("chart-weekly-imp", "figure"),
        Output("table-weekly-container", "children"),
    ],
    Input("date-picker-range", "start_date"),
)
def update_comparisons(_):
  if df_grafico.empty:
    return go.Figure(), go.Figure(), "", go.Figure(), go.Figure(), ""

  df_calc = df_grafico.copy()
  df_calc["Mes"] = df_calc["Fecha"].dt.strftime("%Y-%m")
  df_calc["Semana_Inicio"] = (
      df_calc["Fecha"]
      .dt.to_period("W-SUN")
      .apply(lambda r: r.start_time.strftime("%Y-%m-%d"))
  )

  # Agrupación Mensual
  monthly = (
      df_calc.groupby("Mes")
      .agg(
          Clics=("Clics", "sum"),
          Impresiones=("Impresiones", "sum"),
          Posicion=("Posición", "mean"),
      )
      .reset_index()
  )
  monthly["CTR%"] = (monthly["Clics"] / monthly["Impresiones"] * 100).round(2)
  monthly["Var Clics %"] = (monthly["Clics"].pct_change() * 100).round(1)
  monthly["Var Imp %"] = (monthly["Impresiones"].pct_change() * 100).round(1)
  monthly["Posición Media"] = monthly["Posicion"].round(1)

  fig_m_clics = px.bar(
      monthly,
      x="Mes",
      y="Clics",
      title="Clics por Mes (MoM)",
      color="Clics",
      color_continuous_scale="Blues",
      text="Clics",
  )
  fig_m_clics.update_traces(textposition="outside")
  fig_m_clics.update_layout(
      template="plotly_white", height=280, margin=dict(t=40, b=20, l=10, r=10)
  )

  fig_m_imp = px.bar(
      monthly,
      x="Mes",
      y="Impresiones",
      title="Impresiones por Mes (MoM)",
      color_discrete_sequence=["#0284C7"],
      text="Impresiones",
  )
  fig_m_imp.update_traces(textposition="outside")
  fig_m_imp.update_layout(
      template="plotly_white", height=280, margin=dict(t=40, b=20, l=10, r=10)
  )

  table_m = dash_table.DataTable(
      columns=[
          {"name": c, "id": c}
          for c in [
              "Mes",
              "Clics",
              "Var Clics %",
              "Impresiones",
              "Var Imp %",
              "CTR%",
              "Posición Media",
          ]
      ],
      data=monthly.to_dict("records"),
      style_header={
          "backgroundColor": "#1E293B",
          "color": "white",
          "fontWeight": "bold",
      },
      style_table={"overflowX": "auto"},
      style_cell={
          "textAlign": "center",
          "fontFamily": "sans-serif",
          "fontSize": "12px",
      },
      page_size=6,
  )

  # Agrupación Semanal
  weekly = (
      df_calc.groupby("Semana_Inicio")
      .agg(
          Clics=("Clics", "sum"),
          Impresiones=("Impresiones", "sum"),
          Posicion=("Posición", "mean"),
      )
      .reset_index()
  )
  weekly["CTR%"] = (weekly["Clics"] / weekly["Impresiones"] * 100).round(2)
  weekly["Var Clics %"] = (weekly["Clics"].pct_change() * 100).round(1)
  weekly["Posición Media"] = weekly["Posicion"].round(1)

  fig_w_clics = px.bar(
      weekly,
      x="Semana_Inicio",
      y="Clics",
      title="Clics por Semana (WoW)",
      color="Clics",
      color_continuous_scale="Teal",
      text="Clics",
  )
  fig_w_clics.update_traces(textposition="outside")
  fig_w_clics.update_layout(
      template="plotly_white", height=280, margin=dict(t=40, b=20, l=10, r=10)
  )

  fig_w_imp = px.bar(
      weekly,
      x="Semana_Inicio",
      y="Impresiones",
      title="Impresiones por Semana (WoW)",
      color_discrete_sequence=["#0D9488"],
      text="Impresiones",
  )
  fig_w_imp.update_traces(textposition="outside")
  fig_w_imp.update_layout(
      template="plotly_white", height=280, margin=dict(t=40, b=20, l=10, r=10)
  )

  table_w = dash_table.DataTable(
      columns=[
          {"name": c, "id": c}
          for c in [
              "Semana_Inicio",
              "Clics",
              "Var Clics %",
              "Impresiones",
              "CTR%",
              "Posición Media",
          ]
      ],
      data=weekly.to_dict("records"),
      style_header={
          "backgroundColor": "#0F766E",
          "color": "white",
          "fontWeight": "bold",
      },
      style_table={"overflowX": "auto"},
      style_cell={
          "textAlign": "center",
          "fontFamily": "sans-serif",
          "fontSize": "12px",
      },
      page_size=6,
  )

  return fig_m_clics, fig_m_imp, table_m, fig_w_clics, fig_w_imp, table_w


# =====================================================================
# CALLBACK 3: CONSULTAS / KEYWORDS
# =====================================================================
@app.callback(
    [
        Output("chart-kw-clics", "figure"),
        Output("chart-kw-imp", "figure"),
        Output("table-kw-container", "children"),
    ],
    [Input("input-kw-search", "value"), Input("top-n-slider", "value")],
)
def update_keywords(search_query, top_n):
  if df_consultas.empty:
    return go.Figure(), go.Figure(), ""

  dff = df_consultas.copy()
  if search_query:
    dff = dff[
        dff["Consultas principales"].str.contains(
            search_query, case=False, na=False
        )
    ]

  top_clics = dff.sort_values("Clics", ascending=False).head(top_n)
  fig_c = px.bar(
      top_clics,
      x="Clics",
      y="Consultas principales",
      orientation="h",
      title=f"Top {top_n} Consultas por Clics",
      color="Clics",
      color_continuous_scale="Purples",
      text="Clics",
  )
  fig_c.update_layout(
      yaxis=dict(autorange="reversed"),
      template="plotly_white",
      height=380,
      margin=dict(l=10, r=20, t=40, b=20),
  )

  top_imp = dff.sort_values("Impresiones", ascending=False).head(top_n)
  fig_i = px.bar(
      top_imp,
      x="Impresiones",
      y="Consultas principales",
      orientation="h",
      title=f"Top {top_n} Consultas por Impresiones (Potencial SEO)",
      color="Impresiones",
      color_continuous_scale="Oranges",
      text="Impresiones",
  )
  fig_i.update_layout(
      yaxis=dict(autorange="reversed"),
      template="plotly_white",
      height=380,
      margin=dict(l=10, r=20, t=40, b=20),
  )

  dff_display = dff.copy()
  dff_display["CTR"] = (dff_display["CTR"] * 100).round(2).astype(str) + "%"
  dff_display["Posición"] = dff_display["Posición"].round(1)

  table = dash_table.DataTable(
      columns=[{"name": col, "id": col} for col in dff_display.columns],
      data=dff_display.to_dict("records"),
      style_header={
          "backgroundColor": "#1E293B",
          "color": "white",
          "fontWeight": "bold",
      },
      style_table={"overflowX": "auto"},
      style_cell={
          "textAlign": "left",
          "fontFamily": "sans-serif",
          "fontSize": "12px",
      },
      page_size=10,
      sort_action="native",
  )

  return fig_c, fig_i, table


# =====================================================================
# CALLBACK 4: PÁGINAS PRINCIPALES
# =====================================================================
@app.callback(
    [
        Output("chart-pages-clics", "figure"),
        Output("chart-pages-imp", "figure"),
        Output("table-pages-container", "children"),
    ],
    Input("top-n-slider", "value"),
)
def update_pages(top_n):
  if df_paginas.empty:
    return go.Figure(), go.Figure(), ""

  dff = df_paginas.copy()
  dff["URL_Corta"] = dff["Páginas principales"].apply(
      lambda x: (
          "/" + "/".join(str(x).split("/")[3:])
          if len(str(x).split("/")) > 3
          else str(x)
      )
  )

  top_c = dff.sort_values("Clics", ascending=False).head(top_n)
  fig_c = px.bar(
      top_c,
      x="Clics",
      y="URL_Corta",
      orientation="h",
      title=f"Top {top_n} Páginas por Clics",
      color="Clics",
      color_continuous_scale="Darkmint",
      text="Clics",
  )
  fig_c.update_layout(
      yaxis=dict(autorange="reversed"),
      template="plotly_white",
      height=380,
      margin=dict(l=10, r=20, t=40, b=20),
  )

  top_i = dff.sort_values("Impresiones", ascending=False).head(top_n)
  fig_i = px.bar(
      top_i,
      x="Impresiones",
      y="URL_Corta",
      orientation="h",
      title=f"Top {top_n} Páginas por Impresiones",
      color="Impresiones",
      color_continuous_scale="Burg",
      text="Impresiones",
  )
  fig_i.update_layout(
      yaxis=dict(autorange="reversed"),
      template="plotly_white",
      height=380,
      margin=dict(l=10, r=20, t=40, b=20),
  )

  disp_pg = df_paginas.copy()
  disp_pg["CTR"] = (disp_pg["CTR"] * 100).round(2).astype(str) + "%"
  disp_pg["Posición"] = disp_pg["Posición"].round(1)

  table = dash_table.DataTable(
      columns=[{"name": col, "id": col} for col in disp_pg.columns],
      data=disp_pg.to_dict("records"),
      style_header={
          "backgroundColor": "#1E293B",
          "color": "white",
          "fontWeight": "bold",
      },
      style_table={"overflowX": "auto"},
      style_cell={
          "textAlign": "left",
          "fontFamily": "sans-serif",
          "fontSize": "12px",
      },
      page_size=10,
      sort_action="native",
  )

  return fig_c, fig_i, table


# =====================================================================
# CALLBACK 5: PAÍSES & DISPOSITIVOS
# =====================================================================
@app.callback(
    [
        Output("chart-countries", "figure"),
        Output("chart-devices", "figure"),
        Output("table-devices-container", "children"),
    ],
    Input("top-n-slider", "value"),
)
def update_geo_devices(top_n):
  if not df_paises.empty:
    top_p = df_paises.sort_values("Clics", ascending=False).head(top_n)
    fig_p = px.bar(
        top_p,
        x="Clics",
        y="País",
        orientation="h",
        title=f"Top {top_n} Países por Clics",
        color="Clics",
        color_continuous_scale="Sunset",
        text="Clics",
    )
    fig_p.update_layout(
        yaxis=dict(autorange="reversed"),
        template="plotly_white",
        height=400,
        margin=dict(l=10, r=20, t=40, b=20),
    )
  else:
    fig_p = go.Figure()

  if not df_dispositivos.empty:
    fig_d = px.bar(
        df_dispositivos,
        x="Dispositivo",
        y="Clics",
        title="Clics por Dispositivo",
        color="Dispositivo",
        color_discrete_map={
            "Móviles": "#2563EB",
            "Ordenador": "#10B981",
            "Tablet": "#F59E0B",
        },
        text="Clics",
    )
    fig_d.update_traces(textposition="outside")
    fig_d.update_layout(
        template="plotly_white",
        height=280,
        showlegend=False,
        margin=dict(t=40, b=20, l=10, r=10),
    )

    disp_dev = df_dispositivos.copy()
    disp_dev["CTR"] = (disp_dev["CTR"] * 100).round(2).astype(str) + "%"
    disp_dev["Posición"] = disp_dev["Posición"].round(1)

    table_d = dash_table.DataTable(
        columns=[{"name": c, "id": c} for c in disp_dev.columns],
        data=disp_dev.to_dict("records"),
        style_header={
            "backgroundColor": "#1E293B",
            "color": "white",
            "fontWeight": "bold",
        },
        style_cell={
            "textAlign": "center",
            "fontFamily": "sans-serif",
            "fontSize": "12px",
        },
        style_table={"overflowX": "auto"},
    )
  else:
    fig_d = go.Figure()
    table_d = ""

  return fig_p, fig_d, table_d


if __name__ == "__main__":
  app.run(debug=True, host="127.0.0.1", port=8050)