"""
dashboard.py
────────────────────────────────────────────────────────────────
Interactive BI Dashboard — Sales Analytics
Author : Jatin Prasad
Purpose: A fully interactive browser-based BI dashboard built
         with Plotly Dash. Filters update all charts in real time.

Features:
  - KPI cards (Revenue, Orders, AOV, Return Rate)
  - Interactive filters: Date Range, Region, Category, Segment
  - Revenue trend line chart (monthly)
  - Revenue by product bar chart
  - Regional breakdown pie chart
  - Sales Rep leaderboard bar chart
  - Order status donut chart
  - All charts update dynamically on filter change

Run:
  python dashboard.py
  Then open http://127.0.0.1:8050 in your browser
────────────────────────────────────────────────────────────────
"""

import pandas as pd
import numpy as np
from dash import Dash, dcc, html, Input, Output, callback
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# ── Load & Prep Data ──────────────────────────────────────────
df = pd.read_csv("dashboard_data.csv")
df["Date"]     = pd.to_datetime(df["Date"])
df["Revenue"]  = pd.to_numeric(df["Revenue"], errors="coerce")
df["Month"]    = df["Date"].dt.to_period("M").dt.to_timestamp()
df["Month_Str"]= df["Date"].dt.strftime("%b %Y")

MIN_DATE = df["Date"].min().date()
MAX_DATE = df["Date"].max().date()

REGIONS    = sorted(df["Region"].unique())
CATEGORIES = sorted(df["Category"].unique())
SEGMENTS   = sorted(df["Segment"].unique())

# ── Colour Palette ────────────────────────────────────────────
COLORS = {
    "bg":        "#0F172A",
    "card":      "#1E293B",
    "border":    "#334155",
    "text":      "#F1F5F9",
    "subtext":   "#94A3B8",
    "accent1":   "#38BDF8",
    "accent2":   "#34D399",
    "accent3":   "#F59E0B",
    "accent4":   "#F87171",
    "chart_bg":  "#1E293B",
}

CHART_PALETTE = ["#38BDF8","#34D399","#F59E0B","#F87171","#A78BFA","#FB923C","#4ADE80","#60A5FA"]

def chart_layout(title):
    return dict(
        title=dict(text=title, font=dict(color=COLORS["text"], size=14), x=0.01),
        paper_bgcolor=COLORS["chart_bg"],
        plot_bgcolor=COLORS["chart_bg"],
        font=dict(color=COLORS["subtext"], family="Inter, Calibri, sans-serif", size=11),
        margin=dict(l=16, r=16, t=44, b=16),
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color=COLORS["subtext"])),
        xaxis=dict(gridcolor="#2D3748", zerolinecolor="#2D3748"),
        yaxis=dict(gridcolor="#2D3748", zerolinecolor="#2D3748"),
    )

# ── App Init ──────────────────────────────────────────────────
app = Dash(__name__, title="Sales BI Dashboard")

# ── Layout ────────────────────────────────────────────────────
app.layout = html.Div(style={"backgroundColor": COLORS["bg"], "minHeight": "100vh",
                              "fontFamily": "Inter, Calibri, sans-serif", "padding": "0"}, children=[

    # ── Top Bar ───────────────────────────────────────────────
    html.Div(style={"backgroundColor": COLORS["card"],
                    "borderBottom": f"1px solid {COLORS['border']}",
                    "padding": "18px 32px", "display": "flex",
                    "alignItems": "center", "justifyContent": "space-between"}, children=[
        html.Div([
            html.H1("📊 Sales BI Dashboard",
                    style={"margin": 0, "color": COLORS["text"], "fontSize": "22px", "fontWeight": "700"}),
            html.P("Interactive Business Intelligence — Ira A. Fulton Schools of Engineering",
                   style={"margin": "2px 0 0", "color": COLORS["subtext"], "fontSize": "12px"}),
        ]),
        html.Div(id="last-updated",
                 style={"color": COLORS["subtext"], "fontSize": "12px"})
    ]),

    # ── Filters ───────────────────────────────────────────────
    html.Div(style={"backgroundColor": COLORS["card"], "padding": "16px 32px",
                    "borderBottom": f"1px solid {COLORS['border']}",
                    "display": "flex", "gap": "24px", "flexWrap": "wrap",
                    "alignItems": "flex-end"}, children=[

        html.Div([
            html.Label("Date Range", style={"color": COLORS["subtext"], "fontSize": "11px",
                                             "fontWeight": "600", "marginBottom": "6px", "display": "block"}),
            dcc.DatePickerRange(
                id="date-filter",
                min_date_allowed=MIN_DATE,
                max_date_allowed=MAX_DATE,
                start_date=MIN_DATE,
                end_date=MAX_DATE,
                display_format="MMM D, YYYY",
                style={"fontSize": "12px"}
            )
        ]),

        html.Div([
            html.Label("Region", style={"color": COLORS["subtext"], "fontSize": "11px",
                                         "fontWeight": "600", "marginBottom": "6px", "display": "block"}),
            dcc.Dropdown(
                id="region-filter",
                options=[{"label": r, "value": r} for r in REGIONS],
                multi=True, placeholder="All Regions",
                style={"width": "200px", "fontSize": "12px"},
            )
        ]),

        html.Div([
            html.Label("Category", style={"color": COLORS["subtext"], "fontSize": "11px",
                                           "fontWeight": "600", "marginBottom": "6px", "display": "block"}),
            dcc.Dropdown(
                id="category-filter",
                options=[{"label": c, "value": c} for c in CATEGORIES],
                multi=True, placeholder="All Categories",
                style={"width": "200px", "fontSize": "12px"},
            )
        ]),

        html.Div([
            html.Label("Segment", style={"color": COLORS["subtext"], "fontSize": "11px",
                                          "fontWeight": "600", "marginBottom": "6px", "display": "block"}),
            dcc.Dropdown(
                id="segment-filter",
                options=[{"label": s, "value": s} for s in SEGMENTS],
                multi=True, placeholder="All Segments",
                style={"width": "180px", "fontSize": "12px"},
            )
        ]),

        html.Div([
            html.Label("Order Status", style={"color": COLORS["subtext"], "fontSize": "11px",
                                               "fontWeight": "600", "marginBottom": "6px", "display": "block"}),
            dcc.Dropdown(
                id="status-filter",
                options=[{"label": s, "value": s} for s in sorted(df["Status"].unique())],
                multi=True, placeholder="All Statuses",
                style={"width": "180px", "fontSize": "12px"},
            )
        ]),
    ]),

    # ── KPI Cards ─────────────────────────────────────────────
    html.Div(id="kpi-cards", style={"display": "flex", "gap": "16px",
                                     "padding": "20px 32px", "flexWrap": "wrap"}),

    # ── Charts Row 1 ──────────────────────────────────────────
    html.Div(style={"display": "grid", "gridTemplateColumns": "2fr 1fr",
                    "gap": "16px", "padding": "0 32px 16px"}, children=[
        html.Div(dcc.Graph(id="revenue-trend",  config={"displayModeBar": False}),
                 style={"backgroundColor": COLORS["card"], "borderRadius": "10px",
                        "border": f"1px solid {COLORS['border']}"}),
        html.Div(dcc.Graph(id="status-donut",   config={"displayModeBar": False}),
                 style={"backgroundColor": COLORS["card"], "borderRadius": "10px",
                        "border": f"1px solid {COLORS['border']}"}),
    ]),

    # ── Charts Row 2 ──────────────────────────────────────────
    html.Div(style={"display": "grid", "gridTemplateColumns": "1fr 1fr 1fr",
                    "gap": "16px", "padding": "0 32px 32px"}, children=[
        html.Div(dcc.Graph(id="product-bar",    config={"displayModeBar": False}),
                 style={"backgroundColor": COLORS["card"], "borderRadius": "10px",
                        "border": f"1px solid {COLORS['border']}"}),
        html.Div(dcc.Graph(id="region-pie",     config={"displayModeBar": False}),
                 style={"backgroundColor": COLORS["card"], "borderRadius": "10px",
                        "border": f"1px solid {COLORS['border']}"}),
        html.Div(dcc.Graph(id="rep-leaderboard",config={"displayModeBar": False}),
                 style={"backgroundColor": COLORS["card"], "borderRadius": "10px",
                        "border": f"1px solid {COLORS['border']}"}),
    ]),
])

# ── Callback ──────────────────────────────────────────────────
@app.callback(
    Output("kpi-cards",      "children"),
    Output("revenue-trend",  "figure"),
    Output("status-donut",   "figure"),
    Output("product-bar",    "figure"),
    Output("region-pie",     "figure"),
    Output("rep-leaderboard","figure"),
    Output("last-updated",   "children"),
    Input("date-filter",     "start_date"),
    Input("date-filter",     "end_date"),
    Input("region-filter",   "value"),
    Input("category-filter", "value"),
    Input("segment-filter",  "value"),
    Input("status-filter",   "value"),
)
def update_dashboard(start_date, end_date, regions, categories, segments, statuses):
    # ── Filter ────────────────────────────────────────────────
    filtered = df.copy()
    if start_date:
        filtered = filtered[filtered["Date"] >= pd.to_datetime(start_date)]
    if end_date:
        filtered = filtered[filtered["Date"] <= pd.to_datetime(end_date)]
    if regions:
        filtered = filtered[filtered["Region"].isin(regions)]
    if categories:
        filtered = filtered[filtered["Category"].isin(categories)]
    if segments:
        filtered = filtered[filtered["Segment"].isin(segments)]
    if statuses:
        filtered = filtered[filtered["Status"].isin(statuses)]

    completed = filtered[filtered["Status"] == "Completed"]

    # ── KPIs ──────────────────────────────────────────────────
    total_revenue = completed["Revenue"].sum()
    total_orders  = len(filtered)
    aov           = completed["Revenue"].mean() if len(completed) else 0
    return_rate   = len(filtered[filtered["Status"]=="Returned"]) / len(filtered) * 100 if len(filtered) else 0

    def kpi_card(label, value, color, icon):
        return html.Div(style={
            "backgroundColor": COLORS["card"], "borderRadius": "10px",
            "padding": "20px 24px", "flex": "1", "minWidth": "180px",
            "border": f"1px solid {COLORS['border']}",
            "borderTop": f"3px solid {color}"
        }, children=[
            html.P(f"{icon}  {label}", style={"color": COLORS["subtext"], "fontSize": "12px",
                                               "fontWeight": "600", "margin": "0 0 8px"}),
            html.H2(value, style={"color": color, "margin": 0, "fontSize": "26px", "fontWeight": "800"}),
        ])

    kpi_cards = [
        kpi_card("Total Revenue",   f"${total_revenue:,.0f}",     COLORS["accent2"], "💰"),
        kpi_card("Total Orders",    f"{total_orders:,}",           COLORS["accent1"], "📦"),
        kpi_card("Avg Order Value", f"${aov:,.0f}",                COLORS["accent3"], "🧾"),
        kpi_card("Return Rate",     f"{return_rate:.1f}%",         COLORS["accent4"], "↩️"),
        kpi_card("Completed Orders",f"{len(completed):,}",         "#A78BFA",         "✅"),
    ]

    # ── Revenue Trend ─────────────────────────────────────────
    monthly = (completed.groupby("Month")
                         .agg(Revenue=("Revenue","sum"), Orders=("Order_ID","count"))
                         .reset_index())
    monthly["Month_Str"] = monthly["Month"].dt.strftime("%b %Y")
    monthly["Revenue"]   = monthly["Revenue"].round(2)

    fig_trend = go.Figure()
    fig_trend.add_trace(go.Scatter(
        x=monthly["Month_Str"], y=monthly["Revenue"],
        mode="lines+markers", name="Revenue",
        line=dict(color=COLORS["accent2"], width=2.5),
        marker=dict(size=6, color=COLORS["accent2"]),
        fill="tozeroy", fillcolor="rgba(52,211,153,0.08)"
    ))
    fig_trend.add_trace(go.Bar(
        x=monthly["Month_Str"], y=monthly["Orders"],
        name="Orders", yaxis="y2",
        marker_color="rgba(56,189,248,0.25)",
        marker_line_color=COLORS["accent1"],
        marker_line_width=1
    ))
    fig_trend.update_layout(
        **chart_layout("Monthly Revenue Trend"),
        yaxis=dict(title="Revenue ($)", gridcolor="#2D3748", zerolinecolor="#2D3748"),
        yaxis2=dict(title="Orders", overlaying="y", side="right", gridcolor="rgba(0,0,0,0)"),
        legend=dict(orientation="h", y=1.12, x=0.01),
        height=300,
    )

    # ── Status Donut ──────────────────────────────────────────
    status_counts = filtered["Status"].value_counts().reset_index()
    status_counts.columns = ["Status","Count"]
    fig_donut = go.Figure(go.Pie(
        labels=status_counts["Status"], values=status_counts["Count"],
        hole=0.55, marker_colors=CHART_PALETTE,
        textfont=dict(color=COLORS["text"]),
    ))
    fig_donut.update_layout(**chart_layout("Order Status"), height=300,
                             showlegend=True,
                             annotations=[dict(text=f"{total_orders}<br>orders",
                                               x=0.5, y=0.5, showarrow=False,
                                               font=dict(size=14, color=COLORS["text"]))])

    # ── Product Bar ───────────────────────────────────────────
    by_product = (completed.groupby("Product")["Revenue"]
                            .sum().reset_index()
                            .sort_values("Revenue", ascending=True))
    fig_product = go.Figure(go.Bar(
        x=by_product["Revenue"], y=by_product["Product"],
        orientation="h", marker_color=COLORS["accent1"],
        text=by_product["Revenue"].apply(lambda x: f"${x:,.0f}"),
        textposition="outside", textfont=dict(color=COLORS["subtext"], size=10),
    ))
    fig_product.update_layout(**chart_layout("Revenue by Product"), height=300,
                               xaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
                               yaxis=dict(gridcolor="rgba(0,0,0,0)"))

    # ── Region Pie ────────────────────────────────────────────
    by_region = completed.groupby("Region")["Revenue"].sum().reset_index()
    fig_region = go.Figure(go.Pie(
        labels=by_region["Region"], values=by_region["Revenue"].round(2),
        marker_colors=CHART_PALETTE,
        textinfo="label+percent",
        textfont=dict(color=COLORS["text"]),
    ))
    fig_region.update_layout(**chart_layout("Revenue by Region"), height=300)

    # ── Rep Leaderboard ───────────────────────────────────────
    by_rep = (completed.groupby("Sales_Rep")["Revenue"]
                        .sum().reset_index()
                        .sort_values("Revenue", ascending=True)
                        .tail(8))
    bar_colors = [COLORS["accent3"] if i == len(by_rep)-1 else COLORS["accent1"]
                  for i in range(len(by_rep))]
    fig_rep = go.Figure(go.Bar(
        x=by_rep["Revenue"], y=by_rep["Sales_Rep"],
        orientation="h", marker_color=bar_colors,
        text=by_rep["Revenue"].apply(lambda x: f"${x:,.0f}"),
        textposition="outside", textfont=dict(color=COLORS["subtext"], size=10),
    ))
    fig_rep.update_layout(**chart_layout("Sales Rep Leaderboard"), height=300,
                           xaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
                           yaxis=dict(gridcolor="rgba(0,0,0,0)"))

    last_updated = f"Last updated: {datetime.now().strftime('%H:%M:%S')}"

    return kpi_cards, fig_trend, fig_donut, fig_product, fig_region, fig_rep, last_updated


# ── Run ───────────────────────────────────────────────────────
if __name__ == "__main__":
    print("\n" + "═" * 52)
    print("   INTERACTIVE BI DASHBOARD")
    print("═" * 52)
    print("  🌐  Open your browser at: http://127.0.0.1:8050")
    print("  ⌨️   Press Ctrl+C to stop")
    print("═" * 52 + "\n")
    app.run(debug=False, port=8050)
