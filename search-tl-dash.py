import os
import dash
from dash import dcc, html, Input, Output, dash_table
import dash_bootstrap_components as dbc
import plotly.graph_objects as go
import plotly.express as px
import pandas as pd

# -------------------------------------------------------------
# 1. CARGA DE DATOS
# -------------------------------------------------------------
excel_path = "search-consoletl.xlsx"

xls = pd.ExcelFile(excel_path)
df_g = pd.read_excel(xls, sheet_name='Gráfico')
df_c = pd.read_excel(xls, sheet_name='Consultas')
df_p = pd.read_excel(xls, sheet_name='Páginas')
df_pa = pd.read_excel(xls, sheet_name='Países')
df_d = pd.read_excel(xls, sheet_name='Dispositivos')
df_f = pd.read_excel(xls, sheet_name='Filtros')

# Formatear Gráfico
df_g['Fecha'] = pd.to_datetime(df_g['Fecha'])
df_g = df_g.sort_values('Fecha').reset_index(drop=True)

min_date = df_g['Fecha'].min().date()
max_date = df_g['Fecha'].max().date()

# -------------------------------------------------------------
# 2. CONFIGURACIÓN VISUAL (TEMA E2VISUAL)
# -------------------------------------------------------------
COLOR_DARK = "#0a0c10"
COLOR_CARD = "#12161f"
COLOR_PRIMARY = "#00d2ff"
COLOR_SECONDARY = "#3a7bd5"
COLOR_ACCENT = "#7928ca"
COLOR_TEXT_MUTED = "#8b949e"
COLOR_BORDER = "rgba(255, 255, 255, 0.08)"

card_style = {
    "backgroundColor": COLOR_CARD,
    "border": f"1px solid {COLOR_BORDER}",
    "borderRadius": "14px",
    "padding": "1.25rem",
    "boxShadow": "0 8px 24px rgba(0,0,0,0.2)"
}

table_header_style = {
    'backgroundColor': 'rgba(255, 255, 255, 0.04)',
    'fontWeight': '600',
    'color': '#ffffff',
    'borderBottom': f'1px solid {COLOR_BORDER}',
    'fontSize': '0.78rem',
    'textTransform': 'uppercase',
    'letterSpacing': '0.05em'
}

table_cell_style = {
    'backgroundColor': COLOR_CARD,
    'color': '#c9d1d9',
    'borderBottom': f'1px solid {COLOR_BORDER}',
    'fontSize': '0.84rem',
    'padding': '8px 12px',
    'whiteSpace': 'normal',
    'height': 'auto',
}

# -------------------------------------------------------------
# 3. INICIALIZAR APP DASH
# -------------------------------------------------------------
app = dash.Dash(
    __name__,
    external_stylesheets=[
        dbc.themes.DARKLY,
        "https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css",
        "https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Space+Grotesk:wght@500;600;700&display=swap"
    ],
    title="SEO Analytics Hub • e2visual"
)
server = app.server
app = server

# Badges de filtros del archivo original
filter_badges = [
    html.Span(
        [html.I(className="bi bi-funnel text-info me-1"), f"{r['Filtrar']}: ", html.B(str(r['Valor']), className="text-white")],
        className="badge me-2 py-2 px-3",
        style={"backgroundColor": "rgba(255,255,255,0.05)", "border": f"1px solid {COLOR_BORDER}"}
    )
    for _, r in df_f.iterrows()
]

# -------------------------------------------------------------
# 4. LAYOUT
# -------------------------------------------------------------
app.layout = html.Div(
    style={"backgroundColor": COLOR_DARK, "minHeight": "100vh", "fontFamily": "'Inter', sans-serif"},
    children=[
        # Navbar
        dbc.Navbar(
            dbc.Container(
                [
                    html.Div(
                        [
                            html.Span("e2", className="fs-3 fw-bold text-white", style={"fontFamily": "'Space Grotesk', sans-serif"}),
                            html.Span("visual", className="fs-3 fw-bold", style={"color": COLOR_PRIMARY, "fontFamily": "'Space Grotesk', sans-serif"}),
                            dbc.Badge("SEO Analytics Hub", className="ms-3 d-none d-sm-inline-block", style={"backgroundColor": "rgba(0,210,255,0.15)", "color": COLOR_PRIMARY, "border": f"1px solid {COLOR_PRIMARY}"})
                        ],
                        className="d-flex align-items-center"
                    ),
                    html.Div(
                        [
                            html.Small("Propiedad analizada", className="text-muted d-block text-end"),
                            html.Strong("tradinglatinoacademy.com", className="text-white small")
                        ]
                    )
                ],
                fluid=True,
                className="px-lg-5"
            ),
            color="#0d1117",
            className="border-bottom border-secondary py-3 mb-4",
            style={"backdropFilter": "blur(12px)"}
        ),

        # Contenedor Principal
        dbc.Container(
            fluid=True,
            className="px-lg-5 pb-5",
            children=[
                # Cabecera con selector interactivo de fechas
                dbc.Row(
                    [
                        dbc.Col(
                            [
                                html.H2(["Rendimiento Orgánico ", html.Span("en Buscadores", style={"color": COLOR_PRIMARY, "fontFamily": "'Space Grotesk', sans-serif"})], className="fw-bold mb-1"),
                                html.P("Auditoría y análisis cuantitativo de tráfico web, CTR, impresiones y posicionamiento SEO.", className="text-muted")
                            ],
                            lg=7
                        ),
                        dbc.Col(
                            [
                                html.Div(
                                    [
                                        html.Label("Rango de fechas interactivo:", className="small text-muted d-block mb-1"),
                                        dcc.DatePickerRange(
                                            id='date-picker-range',
                                            min_date_allowed=min_date,
                                            max_date_allowed=max_date,
                                            start_date=min_date,
                                            end_date=max_date,
                                            display_format='YYYY-MM-DD',
                                            style={"borderRadius": "8px", "overflow": "hidden"}
                                        )
                                    ],
                                    className="d-flex flex-column align-items-lg-end"
                                )
                            ],
                            lg=5
                        )
                    ],
                    className="align-items-center mb-3 g-3"
                ),

                # Filtros aplicados
                html.Div(filter_badges, className="mb-4 d-flex flex-wrap gap-2"),

                # KPIs Dinámicos
                html.Div(id='kpi-container', className="mb-4"),

                # Gráfico Principal y Tabs de Frecuencia
                dbc.Row(
                    [
                        dbc.Col(
                            html.Div(
                                [
                                    dbc.Tabs(
                                        [
                                            dbc.Tab(label="Evolución Diaria", tab_id="tab-daily"),
                                            dbc.Tab(label="Comparativa Semanal", tab_id="tab-weekly"),
                                            dbc.Tab(label="Comparativa Mensual", tab_id="tab-monthly"),
                                        ],
                                        id="granularity-tabs",
                                        active_tab="tab-daily",
                                        className="mb-3"
                                    ),
                                    dcc.Graph(id='main-trend-chart', config={"displayModeBar": False})
                                ],
                                style=card_style
                            ),
                            width=12
                        )
                    ],
                    className="mb-4"
                ),

                # Top Consultas y Top Páginas
                dbc.Row(
                    [
                        dbc.Col(
                            html.Div(
                                [
                                    html.H5([html.I(className="bi bi-search text-info me-2"), "Top Consultas (Keywords)"], className="fw-bold text-white mb-3"),
                                    dcc.Graph(id='keywords-bar-chart', config={"displayModeBar": False}, style={"height": "260px"}),
                                    dash_table.DataTable(
                                        id='keywords-table',
                                        columns=[
                                            {"name": "Consulta", "id": "Consultas principales"},
                                            {"name": "Clics", "id": "Clics", "type": "numeric"},
                                            {"name": "Imp.", "id": "Impresiones", "type": "numeric"},
                                            {"name": "CTR %", "id": "CTR_pct"},
                                            {"name": "Posición", "id": "Posición"}
                                        ],
                                        page_size=8,
                                        style_header=table_header_style,
                                        style_cell=table_cell_style,
                                        style_table={"overflowX": "auto"}
                                    )
                                ],
                                style=card_style
                            ),
                            lg=6,
                            className="mb-4"
                        ),
                        dbc.Col(
                            html.Div(
                                [
                                    html.H5([html.I(className="bi bi-link-45deg text-primary me-2"), "Páginas Principales"], className="fw-bold text-white mb-3"),
                                    dcc.Graph(id='pages-bar-chart', config={"displayModeBar": False}, style={"height": "260px"}),
                                    dash_table.DataTable(
                                        id='pages-table',
                                        columns=[
                                            {"name": "Página", "id": "Páginas principales"},
                                            {"name": "Clics", "id": "Clics", "type": "numeric"},
                                            {"name": "Imp.", "id": "Impresiones", "type": "numeric"},
                                            {"name": "CTR %", "id": "CTR_pct"},
                                            {"name": "Posición", "id": "Posición"}
                                        ],
                                        page_size=8,
                                        style_header=table_header_style,
                                        style_cell=table_cell_style,
                                        style_table={"overflowX": "auto"}
                                    )
                                ],
                                style=card_style
                            ),
                            lg=6,
                            className="mb-4"
                        )
                    ]
                ),

                # Países y Dispositivos
                dbc.Row(
                    [
                        dbc.Col(
                            html.Div(
                                [
                                    html.H5([html.I(className="bi bi-globe-americas text-info me-2"), "Distribución Geográfica"], className="fw-bold text-white mb-3"),
                                    dcc.Graph(id='countries-bar-chart', config={"displayModeBar": False}, style={"height": "280px"}),
                                    dash_table.DataTable(
                                        id='countries-table',
                                        columns=[
                                            {"name": "País", "id": "País"},
                                            {"name": "Clics", "id": "Clics"},
                                            {"name": "Impresiones", "id": "Impresiones"},
                                            {"name": "CTR %", "id": "CTR_pct"},
                                            {"name": "Posición", "id": "Posición"}
                                        ],
                                        page_size=6,
                                        style_header=table_header_style,
                                        style_cell=table_cell_style
                                    )
                                ],
                                style=card_style
                            ),
                            lg=7,
                            className="mb-4"
                        ),
                        dbc.Col(
                            html.Div(
                                [
                                    html.H5([html.I(className="bi bi-devices text-primary me-2"), "Segmentación por Dispositivo"], className="fw-bold text-white mb-3"),
                                    dcc.Graph(id='devices-pie-chart', config={"displayModeBar": False}, style={"height": "280px"}),
                                    dash_table.DataTable(
                                        id='devices-table',
                                        columns=[
                                            {"name": "Dispositivo", "id": "Dispositivo"},
                                            {"name": "Clics", "id": "Clics"},
                                            {"name": "Imp.", "id": "Impresiones"},
                                            {"name": "CTR %", "id": "CTR_pct"}
                                        ],
                                        style_header=table_header_style,
                                        style_cell=table_cell_style
                                    )
                                ],
                                style=card_style
                            ),
                            lg=5,
                            className="mb-4"
                        )
                    ]
                )
            ]
        )
    ]
)

# -------------------------------------------------------------
# 5. CALLBACKS (LÓGICA INTERACTIVA)
# -------------------------------------------------------------
@app.callback(
    [
        Output('kpi-container', 'children'),
        Output('main-trend-chart', 'figure'),
        Output('keywords-bar-chart', 'figure'),
        Output('keywords-table', 'data'),
        Output('pages-bar-chart', 'figure'),
        Output('pages-table', 'data'),
        Output('countries-bar-chart', 'figure'),
        Output('countries-table', 'data'),
        Output('devices-pie-chart', 'figure'),
        Output('devices-table', 'data'),
    ],
    [
        Input('date-picker-range', 'start_date'),
        Input('date-picker-range', 'end_date'),
        Input('granularity-tabs', 'active_tab')
    ]
)
def update_dashboard(start_date, end_date, active_tab):
    # Filtrar gráfico por rango seleccionado
    mask = (df_g['Fecha'] >= pd.to_datetime(start_date)) & (df_g['Fecha'] <= pd.to_datetime(end_date))
    dff = df_g.loc[mask].copy()

    if dff.empty:
        dff = df_g.copy()

    total_clics = int(dff['Clics'].sum())
    total_imp = int(dff['Impresiones'].sum())
    ctr_pond = round((total_clics / total_imp) * 100, 2) if total_imp > 0 else 0
    pos_media = round(float(dff['Posición'].mean()), 1)

    # Mitades para deltas
    total_dias = len(dff)
    half = total_dias // 2
    if half > 0:
        recent_half = dff.iloc[half:]
        prev_half = dff.iloc[:half]

        clics_rec, clics_prev = recent_half['Clics'].sum(), prev_half['Clics'].sum()
        delta_clics = round(((clics_rec - clics_prev) / clics_prev * 100), 1) if clics_prev > 0 else 0

        imp_rec, imp_prev = recent_half['Impresiones'].sum(), prev_half['Impresiones'].sum()
        delta_imp = round(((imp_rec - imp_prev) / imp_prev * 100), 1) if imp_prev > 0 else 0

        ctr_rec = (clics_rec / imp_rec * 100) if imp_rec > 0 else 0
        ctr_prev = (clics_prev / imp_prev * 100) if imp_prev > 0 else 0
        delta_ctr = round(ctr_rec - ctr_prev, 2)

        pos_rec = recent_half['Posición'].mean()
        pos_prev = prev_half['Posición'].mean()
        delta_pos = round(-(pos_rec - pos_prev), 1)
    else:
        delta_clics = delta_imp = delta_ctr = delta_pos = 0

    # Layout de KPIs
    def make_kpi(title, value, delta, suffix="", delta_suffix="%", icon="bi-cursor-fill", icon_color=COLOR_PRIMARY):
        is_pos = delta >= 0
        badge_bg = "rgba(46, 160, 67, 0.15)" if is_pos else "rgba(248, 81, 73, 0.15)"
        badge_color = "#3fb950" if is_pos else "#f85149"
        arrow = "bi-arrow-up-right" if is_pos else "bi-arrow-down-right"

        return dbc.Col(
            html.Div(
                [
                    html.Div(
                        [
                            html.Span(title, className="small text-muted text-uppercase fw-semibold"),
                            html.I(className=f"bi {icon} fs-5", style={"color": icon_color})
                        ],
                        className="d-flex justify-content-between align-items-center mb-2"
                    ),
                    html.Div(f"{value}{suffix}", className="fs-2 fw-bold text-white", style={"fontFamily": "'Space Grotesk', sans-serif"}),
                    html.Div(
                        [
                            html.Span(
                                [html.I(className=f"bi {arrow} me-1"), f"{'+' if delta > 0 else ''}{delta}{delta_suffix}"],
                                className="badge px-2 py-1",
                                style={"backgroundColor": badge_bg, "color": badge_color}
                            ),
                            html.Span("vs período previo", className="small text-muted ms-2")
                        ],
                        className="mt-2"
                    )
                ],
                style=card_style
            ),
            sm=6,
            xl=3
        )

    kpi_children = dbc.Row(
        [
            make_kpi("Clics Totales", f"{total_clics:,}", delta_clics, icon="bi-cursor-fill", icon_color=COLOR_PRIMARY),
            make_kpi("Impresiones Totales", f"{total_imp:,}", delta_imp, icon="bi-eye-fill", icon_color=COLOR_SECONDARY),
            make_kpi("CTR Medio Ponderado", f"{ctr_pond:.2f}", delta_ctr, suffix="%", delta_suffix=" pts", icon="bi-percent", icon_color="#2ea043"),
            make_kpi("Posición Media", f"{pos_media:.1f}", delta_pos, delta_suffix=" pts", icon="bi-trophy-fill", icon_color="#e3b341"),
        ],
        className="g-3"
    )

    # 1. Gráfico de Tendencia (Diario / Semanal / Mensual)
    fig_main = go.Figure()
    if active_tab == "tab-monthly":
        dff['Mes'] = dff['Fecha'].dt.strftime('%Y-%m')
        agg_data = dff.groupby('Mes').agg(Clics=('Clics', 'sum'), Impresiones=('Impresiones', 'sum')).reset_index()
        x_col = agg_data['Mes']
    elif active_tab == "tab-weekly":
        dff['Semana'] = dff['Fecha'].dt.to_period('W-SUN').apply(lambda r: r.start_time.strftime('%Y-%m-%d'))
        agg_data = dff.groupby('Semana').agg(Clics=('Clics', 'sum'), Impresiones=('Impresiones', 'sum')).reset_index()
        x_col = agg_data['Semana']
    else:
        agg_data = dff
        x_col = agg_data['Fecha']

    fig_main.add_trace(go.Bar(
        x=x_col,
        y=agg_data['Clics'],
        name='Clics',
        marker_color=COLOR_PRIMARY,
        yaxis='y1'
    ))
    fig_main.add_trace(go.Scatter(
        x=x_col,
        y=agg_data['Impresiones'],
        name='Impresiones',
        line=dict(color=COLOR_SECONDARY, width=2.5),
        mode='lines+markers',
        yaxis='y2'
    ))
    fig_main.update_layout(
        template="plotly_dark",
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        height=320,
        margin=dict(l=20, r=20, t=20, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        yaxis=dict(title="Clics", showgrid=True, gridcolor="rgba(255,255,255,0.05)"),
        yaxis2=dict(title="Impresiones", overlaying='y', side='right', showgrid=False),
        xaxis=dict(showgrid=False)
    )

    # 2. Gráfico y Datos de Keywords
    top_c = df_c.head(10).copy()
    fig_kw = px.bar(
        top_c.sort_values('Clics', ascending=True),
        x='Clics',
        y='Consultas principales',
        orientation='h',
        color_discrete_sequence=[COLOR_PRIMARY]
    )
    fig_kw.update_layout(
        template="plotly_dark",
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=10, r=10, t=10, b=10),
        xaxis=dict(showgrid=True, gridcolor="rgba(255,255,255,0.05)"),
        yaxis=dict(showgrid=False, automargin=True)
    )
    c_table = df_c.copy()
    c_table['CTR_pct'] = (c_table['CTR'] * 100).round(2).astype(str) + "%"
    c_table['Posición'] = c_table['Posición'].round(1)

    # 3. Gráfico y Datos de Páginas
    top_p_data = df_p.head(10).copy()
    top_p_data['path'] = top_p_data['Páginas principales'].apply(lambda x: "/" + "/".join(str(x).split("/")[3:]) if len(str(x).split("/")) > 3 else str(x))
    fig_p = px.bar(
        top_p_data.sort_values('Clics', ascending=True),
        x='Clics',
        y='path',
        orientation='h',
        color_discrete_sequence=[COLOR_SECONDARY]
    )
    fig_p.update_layout(
        template="plotly_dark",
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=10, r=10, t=10, b=10),
        xaxis=dict(showgrid=True, gridcolor="rgba(255,255,255,0.05)"),
        yaxis=dict(showgrid=False, automargin=True)
    )
    p_table = df_p.copy()
    p_table['CTR_pct'] = (p_table['CTR'] * 100).round(2).astype(str) + "%"
    p_table['Posición'] = p_table['Posición'].round(1)

    # 4. Gráfico y Datos de Países
    top_pa = df_pa.head(8).copy()
    fig_pa = px.bar(
        top_pa.sort_values('Clics', ascending=True),
        x='Clics',
        y='País',
        orientation='h',
        color_discrete_sequence=[COLOR_PRIMARY]
    )
    fig_pa.update_layout(
        template="plotly_dark",
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=10, r=10, t=10, b=10),
        xaxis=dict(showgrid=True, gridcolor="rgba(255,255,255,0.05)"),
        yaxis=dict(showgrid=False, automargin=True)
    )
    pa_table = df_pa.copy()
    pa_table['CTR_pct'] = (pa_table['CTR'] * 100).round(2).astype(str) + "%"
    pa_table['Posición'] = pa_table['Posición'].round(1)

    # 5. Gráfico y Datos de Dispositivos
    fig_d = px.pie(
        df_d,
        names='Dispositivo',
        values='Clics',
        hole=0.6,
        color_discrete_sequence=[COLOR_PRIMARY, COLOR_SECONDARY, COLOR_ACCENT]
    )
    fig_d.update_layout(
        template="plotly_dark",
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=10, r=10, t=10, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=-0.1, xanchor="center", x=0.5)
    )
    d_table = df_d.copy()
    d_table['CTR_pct'] = (d_table['CTR'] * 100).round(2).astype(str) + "%"

    return (
        kpi_children,
        fig_main,
        fig_kw,
        c_table.to_dict('records'),
        fig_p,
        p_table.to_dict('records'),
        fig_pa,
        pa_table.to_dict('records'),
        fig_d,
        d_table.to_dict('records')
    )

# -------------------------------------------------------------
# 6. EJECUCIÓN DEL SERVIDOR
# -------------------------------------------------------------
if __name__ == '__main__':
    app.run(debug=True, port=8050)
