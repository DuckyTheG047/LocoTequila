from pathlib import Path
from typing import Optional

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

BASE_DIR = Path(__file__).parent
ENGAGEMENT_FILE = BASE_DIR / (
    "EngagementOverview-locotequila.mx_claseazul.com_casadragones.com.mx_"
    "herradura.com_tequilafortaleza.com-(999)-(2026_03-2026_05).xlsx"
)
TWO_SITE_FILE = BASE_DIR / "loco_tequila_web_traffic_database.xlsx"
COMPETITOR_DB_FILE = BASE_DIR / "loco_tequila_competitors_web_traffic_database.xlsx"

PRIMARY_SITE = "locotequila.mx"
SITE_LABELS = {
    "loco-tequila.com": "Loco Tequila .com",
    "locotequila.mx": "Loco Tequila",
    "claseazul.com": "Clase Azul",
    "casadragones.com.mx": "Casa Dragones",
    "herradura.com": "Herradura",
    "tequilafortaleza.com": "Tequila Fortaleza",
}
PALETTE = {
    "loco-tequila.com": "#8a4b18",
    "locotequila.mx": "#c47a2c",
    "claseazul.com": "#0e4c92",
    "casadragones.com.mx": "#7a1f33",
    "herradura.com": "#3d2c1e",
    "tequilafortaleza.com": "#1f6f43",
}


def apply_dashboard_styles() -> None:
    st.markdown(
        """
        <style>
        :root {
            color-scheme: light;
            --lt-ink: #12100e;
            --lt-char: #241c16;
            --lt-earth: #5a4330;
            --lt-amber: #b97830;
            --lt-amber-soft: #d9aa6d;
            --lt-ivory: #f7f1e6;
            --lt-bone: #efe6d6;
            --lt-mist: #fcf8f1;
            --lt-muted: #6b5a4a;
            --lt-line: rgba(101, 76, 49, 0.16);
        }
        html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
            background: #fbf7f1 !important;
            color: var(--lt-ink) !important;
        }
        .stApp {
            background:
                radial-gradient(circle at top left, rgba(185, 120, 48, 0.12), transparent 24%),
                radial-gradient(circle at 85% 0%, rgba(123, 95, 67, 0.10), transparent 22%),
                linear-gradient(180deg, #fbf7f1 0%, #f4ecdf 54%, #fdf9f3 100%);
            color: var(--lt-ink);
        }
        [data-testid="stHeader"] {
            background: rgba(251,247,241,0.88) !important;
        }
        [data-testid="stSidebar"] {
            background: #f7f1e6 !important;
        }
        .block-container {
            padding-top: 1.4rem;
            padding-bottom: 3rem;
            max-width: 1440px;
        }
        h1, h2, h3 {
            color: var(--lt-ink);
            letter-spacing: -0.02em;
            font-family: "Georgia", "Times New Roman", serif;
        }
        p, li, label, [data-testid="stMarkdownContainer"] {
            font-family: "Helvetica Neue", "Arial", sans-serif;
        }
        [data-baseweb="tab-list"] {
            gap: 8px;
            margin: 0 0 14px 0;
        }
        [data-baseweb="tab"] {
            border-radius: 999px;
            background: rgba(255,255,255,0.72);
            border: 1px solid var(--lt-line);
            color: var(--lt-earth);
            padding: 10px 16px;
        }
        [aria-selected="true"][data-baseweb="tab"] {
            background: linear-gradient(180deg, #1a1511 0%, #30241c 100%);
            color: #f8f1e5;
            border-color: rgba(26,21,17,0.1);
        }
        .dashboard-hero {
            background:
                linear-gradient(180deg, rgba(255,255,255,0.75) 0%, rgba(255,255,255,0.62) 100%),
                linear-gradient(135deg, #f6efe2 0%, #fbf8f2 100%);
            border: 1px solid var(--lt-line);
            border-radius: 28px;
            padding: 22px 28px 22px 28px;
            box-shadow: 0 24px 60px rgba(36, 28, 22, 0.08);
            margin-bottom: 18px;
            position: relative;
            overflow: hidden;
        }
        .hero-brand {
            display: flex;
            align-items: center;
            gap: 18px;
            margin: 0 0 10px 0;
        }
        .hero-logo {
            height: 58px;
            width: auto;
            display: block;
            filter: saturate(0.92) contrast(1.02);
        }
        .hero-copy {
            display: flex;
            flex-direction: column;
        }
        .dashboard-hero::before {
            content: "";
            position: absolute;
            inset: 0 auto 0 0;
            width: 10px;
            background: linear-gradient(180deg, #1a1511 0%, var(--lt-amber) 55%, #d8b587 100%);
        }
        .dashboard-hero h1 {
            color: var(--lt-ink);
            font-size: 2.5rem;
            margin: 0 0 6px 0;
        }
        .dashboard-hero p {
            color: var(--lt-muted);
            margin: 0;
            font-size: 0.98rem;
            line-height: 1.6;
            max-width: 760px;
        }
        .hero-kicker {
            display: inline-block;
            font-size: 0.74rem;
            letter-spacing: 0.18em;
            text-transform: uppercase;
            color: var(--lt-amber);
            margin-bottom: 10px;
            padding: 7px 11px 6px 13px;
            border: 1px solid rgba(185,120,48,0.22);
            border-radius: 999px;
            background: rgba(255,255,255,0.7);
        }
        .filters-shell {
            background: rgba(255,255,255,0.72);
            border: 1px solid var(--lt-line);
            backdrop-filter: blur(10px);
            border-radius: 22px;
            padding: 18px 20px 4px 20px;
            box-shadow: 0 12px 36px rgba(36, 28, 22, 0.05);
            margin-bottom: 10px;
        }
        .chart-shell {
            background: linear-gradient(180deg, rgba(255,255,255,0.92) 0%, rgba(252,248,241,0.92) 100%);
            border: 1px solid var(--lt-line);
            border-radius: 24px;
            padding: 18px 18px 10px 18px;
            box-shadow: 0 14px 34px rgba(36, 28, 22, 0.06);
            margin-bottom: 18px;
        }
        .chart-shell.featured {
            background: linear-gradient(180deg, rgba(255,255,255,0.98) 0%, rgba(247,241,230,0.98) 100%);
            border: 1px solid rgba(185,120,48,0.18);
            border-radius: 30px;
            padding: 22px 22px 12px 22px;
            box-shadow: 0 20px 52px rgba(36, 28, 22, 0.08), inset 0 1px 0 rgba(255,255,255,0.78);
            position: relative;
            overflow: clip;
        }
        .chart-shell.featured::before {
            content: "";
            position: absolute;
            inset: 0 0 auto 0;
            height: 5px;
            background: linear-gradient(90deg, #18120e 0%, var(--lt-amber) 52%, var(--lt-amber-soft) 100%);
        }
        .chart-shell h3 {
            margin: 0 0 2px 0;
            font-size: 1.08rem;
            color: var(--lt-ink);
        }
        .chart-shell.featured h3 {
            font-size: 1.26rem;
            margin-top: 4px;
        }
        .chart-shell p {
            margin: 0 0 12px 0;
            color: var(--lt-muted);
            font-size: 0.92rem;
            line-height: 1.45;
        }
        .chart-shell.featured p {
            font-size: 0.96rem;
            max-width: 760px;
            margin-bottom: 16px;
        }
        .metric-card {
            border: 1px solid var(--lt-line);
            border-radius: 22px;
            padding: 18px 18px 16px 18px;
            background: linear-gradient(180deg,#fffdfa 0%,#f7efe2 100%);
            box-shadow: 0 10px 28px rgba(36, 28, 22, 0.05);
            min-height: 160px;
            position: relative;
            overflow: hidden;
        }
        .metric-card::before {
            content: "";
            position: absolute;
            inset: 0 auto 0 0;
            width: 4px;
            background: linear-gradient(180deg, var(--lt-amber) 0%, #e0bc8a 100%);
        }
        .metric-kicker {
            font-size: 0.76rem;
            font-weight: 700;
            letter-spacing: 0.16em;
            text-transform: uppercase;
            color: var(--lt-amber);
            margin-bottom: 10px;
        }
        .metric-headline {
            font-size: 1.05rem;
            font-weight: 700;
            line-height: 1.35;
            color: var(--lt-ink);
            margin-bottom: 10px;
        }
        .metric-copy {
            font-size: 0.93rem;
            line-height: 1.5;
            color: var(--lt-muted);
        }
        .stSelectbox label, .stMultiSelect label, .stDateInput label, .stCheckbox label {
            color: var(--lt-earth) !important;
            font-size: 0.85rem !important;
        }
        .stSelectbox [data-baseweb="select"] > div,
        .stMultiSelect [data-baseweb="select"] > div,
        .stDateInput [data-baseweb="input"] > div {
            border-radius: 16px !important;
            border-color: var(--lt-line) !important;
            background: rgba(255,255,255,0.78) !important;
        }
        .stCheckbox > label, .stRadio > label, .stMarkdown, .stCaption {
            color: var(--lt-ink) !important;
        }
        @media (max-width: 720px) {
            .hero-brand {
                align-items: flex-start;
                flex-direction: column;
                gap: 10px;
            }
            .hero-logo {
                height: 48px;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def open_chart_shell(title: str, subtitle: str, featured: bool = False) -> None:
    shell_class = "chart-shell featured" if featured else "chart-shell"
    st.markdown(
        f"""
        <div class="{shell_class}">
            <h3>{title}</h3>
            <p>{subtitle}</p>
        """,
        unsafe_allow_html=True,
    )


def close_chart_shell() -> None:
    st.markdown("</div>", unsafe_allow_html=True)


def display_site_name(site: str) -> str:
    return SITE_LABELS.get(site, site)


def parse_duration_to_seconds(value: object) -> float:
    if pd.isna(value):
        return 0.0
    text = str(value).strip()
    if not text:
        return 0.0
    parts = text.split(":")
    if len(parts) != 3:
        return 0.0
    hours, minutes, seconds = (int(part) for part in parts)
    return float((hours * 3600) + (minutes * 60) + seconds)


def format_seconds_label(seconds: float) -> str:
    total = int(round(seconds))
    minutes, remaining_seconds = divmod(total, 60)
    hours, remaining_minutes = divmod(minutes, 60)
    if hours > 0:
        return f"{hours:02d}:{remaining_minutes:02d}:{remaining_seconds:02d}"
    return f"{remaining_minutes:02d}:{remaining_seconds:02d}"


def safe_pct_change(current: float, previous: float) -> Optional[float]:
    if previous == 0 or pd.isna(previous):
        return None
    return ((current - previous) / previous) * 100


def format_delta(value: Optional[float]) -> str:
    if value is None:
        return "N/A"
    sign = "+" if value > 0 else ""
    return f"{sign}{value:.1f}%"


@st.cache_data(show_spinner=False)
def load_source_frames() -> dict[str, pd.DataFrame]:
    if not ENGAGEMENT_FILE.exists():
        raise FileNotFoundError(f"No se encontro el archivo {ENGAGEMENT_FILE}")

    workbook = pd.ExcelFile(ENGAGEMENT_FILE)
    source_frames: dict[str, pd.DataFrame] = {}

    for sheet_name in workbook.sheet_names:
        if sheet_name in {"Report Details", "Aggregated_Data_for_Site"}:
            continue
        df = workbook.parse(sheet_name).copy()
        df["Date"] = pd.to_datetime(df["Date"])
        for metric in ["Visits", "Desktop Share", "Pages / Visit", "Bounce Rate", "Total Page Views"]:
            if metric in df.columns:
                df[metric] = pd.to_numeric(df[metric], errors="coerce").fillna(0.0)
        df["Avg. Visit Duration (seconds)"] = df["Avg. Visit Duration"].apply(parse_duration_to_seconds)
        df["Mobile Share"] = 1 - df["Desktop Share"]
        df["Site"] = sheet_name
        source_frames[sheet_name] = df

    return source_frames


@st.cache_data(show_spinner=False)
def load_combined_visits(source_frames: dict[str, pd.DataFrame]) -> pd.DataFrame:
    merged_df: Optional[pd.DataFrame] = None
    for site, frame in source_frames.items():
        site_frame = frame[["Date", "Visits"]].rename(columns={"Visits": site})
        if merged_df is None:
            merged_df = site_frame
        else:
            merged_df = merged_df.merge(site_frame, on="Date", how="outer", validate="one_to_one")
    if merged_df is None:
        return pd.DataFrame()
    return merged_df.sort_values("Date").set_index("Date").fillna(0.0)


@st.cache_data(show_spinner=False)
def load_two_site_data() -> dict[str, pd.DataFrame]:
    if not TWO_SITE_FILE.exists():
        raise FileNotFoundError(f"No se encontro el archivo {TWO_SITE_FILE}")

    workbook = pd.ExcelFile(TWO_SITE_FILE)

    summary_df = workbook.parse("domains_summary").copy()
    summary_df["domain_label"] = summary_df["domain"].apply(display_site_name)
    for column in [
        "total_visits",
        "monthly_visits",
        "monthly_unique_visitors",
        "visits_per_unique_visitor",
        "deduplicated_audience",
        "visit_duration_seconds",
        "pages_per_visit",
        "bounce_rate",
        "page_views",
    ]:
        summary_df[column] = pd.to_numeric(summary_df[column], errors="coerce")

    monthly_df = workbook.parse("monthly_visits").copy()
    monthly_df["month"] = pd.to_datetime(monthly_df["month"])
    monthly_df["domain_label"] = monthly_df["domain"].apply(display_site_name)
    monthly_df["visits"] = pd.to_numeric(monthly_df["visits"], errors="coerce")

    device_df = workbook.parse("device_distribution").copy()
    device_df["domain_label"] = device_df["domain"].apply(display_site_name)
    device_df["share"] = pd.to_numeric(device_df["share"], errors="coerce")

    return {"summary": summary_df, "monthly": monthly_df, "device": device_df}


@st.cache_data(show_spinner=False)
def load_competitor_database_summary() -> pd.DataFrame:
    if not COMPETITOR_DB_FILE.exists():
        raise FileNotFoundError(f"No se encontro el archivo {COMPETITOR_DB_FILE}")

    df = pd.read_excel(COMPETITOR_DB_FILE, sheet_name="domains_summary").copy()
    for column in [
        "total_visits_period",
        "monthly_visits",
        "monthly_unique_visitors",
        "visits_per_unique_visitor",
        "deduplicated_audience",
        "pages_per_visit",
        "bounce_rate",
        "page_views",
        "desktop_share",
        "mobile_share",
    ]:
        if column in df.columns:
            df[column] = pd.to_numeric(df[column], errors="coerce")
    df["visit_duration_seconds"] = df["visit_duration"].apply(parse_duration_to_seconds)
    df["domain_label"] = df["domain"].apply(display_site_name)
    return df


def build_summary_table(filtered_frames: dict[str, pd.DataFrame]) -> pd.DataFrame:
    rows = []
    for site, df in filtered_frames.items():
        if df.empty:
            continue
        visits_total = float(df["Visits"].sum())
        latest = df.sort_values("Date").iloc[-1]
        previous = df.sort_values("Date").iloc[-2] if len(df) > 1 else None
        wow_change = (
            safe_pct_change(float(latest["Visits"]), float(previous["Visits"])) if previous is not None else None
        )
        rows.append(
            {
                "Sitio": display_site_name(site),
                "Site": site,
                "Visits": visits_total,
                "Page Views": float(df["Total Page Views"].sum()),
                "Pages / Visit": float(df["Pages / Visit"].mean()),
                "Bounce Rate (%)": float(df["Bounce Rate"].mean() * 100),
                "Avg. Duration (s)": float(df["Avg. Visit Duration (seconds)"].mean()),
                "Desktop (%)": float(df["Desktop Share"].mean() * 100),
                "Mobile (%)": float(df["Mobile Share"].mean() * 100),
                "Peak Week Visits": float(df["Visits"].max()),
                "WoW Latest (%)": wow_change,
            }
        )
    summary_df = pd.DataFrame(rows)
    if summary_df.empty:
        return summary_df
    return summary_df.sort_values("Visits", ascending=False).reset_index(drop=True)


def build_visits_trend_chart(visits_df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    for site in visits_df.columns:
        fig.add_trace(
            go.Scatter(
                x=visits_df.index,
                y=visits_df[site],
                mode="lines+markers",
                name=display_site_name(site),
                line=dict(color=PALETTE.get(site, "#1a1714"), width=3 if site == PRIMARY_SITE else 2),
                marker=dict(size=7),
                hovertemplate="%{x|%d %b %Y}<br>%{y:,.0f} visits<extra></extra>",
            )
        )
    fig.update_layout(
        height=520,
        margin=dict(t=24, r=24, b=24, l=24),
        xaxis_title="Semana",
        yaxis_title="Visits",
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(247,241,231,0.92)",
    )
    return fig


def build_traffic_share_chart(summary_df: pd.DataFrame) -> go.Figure:
    colors = [PALETTE.get(site, "#1a1714") for site in summary_df["Site"]]
    fig = go.Figure(
        data=go.Pie(
            labels=summary_df["Sitio"],
            values=summary_df["Visits"],
            hole=0.42,
            textinfo="label+percent",
            marker=dict(colors=colors),
            hovertemplate="%{label}<br>Visits: %{value:,.0f}<br>Share: %{percent}<extra></extra>",
        )
    )
    fig.update_layout(
        height=420,
        margin=dict(t=24, r=24, b=24, l=24),
        showlegend=False,
        paper_bgcolor="rgba(0,0,0,0)",
    )
    return fig


def build_target_vs_competitor_chart(
    filtered_frames: dict[str, pd.DataFrame], competitor_site: str
) -> go.Figure:
    target_df = filtered_frames[PRIMARY_SITE].sort_values("Date")
    competitor_df = filtered_frames[competitor_site].sort_values("Date")
    compare_df = pd.DataFrame(
        {
            "Date": target_df["Date"],
            display_site_name(PRIMARY_SITE): target_df["Visits"].to_numpy(),
            display_site_name(competitor_site): competitor_df["Visits"].to_numpy(),
        }
    )

    fig = go.Figure()
    for column, color in [
        (display_site_name(PRIMARY_SITE), PALETTE[PRIMARY_SITE]),
        (display_site_name(competitor_site), PALETTE.get(competitor_site, "#1a1714")),
    ]:
        fig.add_trace(
            go.Scatter(
                x=compare_df["Date"],
                y=compare_df[column],
                mode="lines+markers",
                name=column,
                line=dict(color=color, width=3),
                marker=dict(size=7),
                hovertemplate="%{x|%d %b %Y}<br>%{y:,.0f} visits<extra></extra>",
            )
        )
    fig.update_layout(
        height=420,
        margin=dict(t=24, r=24, b=24, l=24),
        xaxis_title="Semana",
        yaxis_title="Visits",
        hovermode="x unified",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(247,241,231,0.92)",
    )
    return fig


def build_kpi_heatmap(summary_df: pd.DataFrame) -> go.Figure:
    metric_columns = [
        "Pages / Visit",
        "Bounce Rate (%)",
        "Avg. Duration (s)",
        "Desktop (%)",
        "Mobile (%)",
    ]
    text_values = []
    for _, row in summary_df.iterrows():
        text_values.append(
            [
                f"{row['Pages / Visit']:.2f}",
                f"{row['Bounce Rate (%)']:.1f}%",
                format_seconds_label(float(row["Avg. Duration (s)"])),
                f"{row['Desktop (%)']:.1f}%",
                f"{row['Mobile (%)']:.1f}%",
            ]
        )

    fig = go.Figure(
        data=go.Heatmap(
            z=summary_df[metric_columns].to_numpy(),
            x=metric_columns,
            y=summary_df["Sitio"],
            text=text_values,
            texttemplate="%{text}",
            colorscale=[
                [0.0, "#fffaf2"],
                [0.2, "#f1dfc7"],
                [0.4, "#e0bd8f"],
                [0.65, "#cb8f4d"],
                [0.82, "#8e5f2d"],
                [1.0, "#2f241d"],
            ],
            hovertemplate="Sitio: %{y}<br>Metrica: %{x}<br>Valor: %{text}<extra></extra>",
            colorbar=dict(title="Nivel"),
        )
    )
    fig.update_layout(
        height=max(420, 70 * len(summary_df)),
        margin=dict(t=24, r=24, b=24, l=24),
        xaxis_title="KPIs",
        yaxis_title="Sitios",
    )
    return fig


def build_quality_scatter(summary_df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    for _, row in summary_df.iterrows():
        fig.add_trace(
            go.Scatter(
                x=[row["Bounce Rate (%)"]],
                y=[row["Pages / Visit"]],
                mode="markers+text",
                name=row["Sitio"],
                text=[row["Sitio"]],
                textposition="top center",
                marker=dict(
                    size=max(16, row["Visits"] ** 0.18),
                    color=PALETTE.get(row["Site"], "#1a1714"),
                    opacity=0.88,
                    line=dict(color="white", width=1.5),
                ),
                hovertemplate=(
                    f"{row['Sitio']}<br>"
                    "Bounce Rate: %{x:.1f}%<br>"
                    "Pages / Visit: %{y:.2f}<br>"
                    f"Visits: {row['Visits']:,.0f}<extra></extra>"
                ),
            )
        )
    fig.update_layout(
        height=420,
        margin=dict(t=24, r=24, b=24, l=24),
        xaxis_title="Bounce Rate (%)",
        yaxis_title="Pages / Visit",
        showlegend=False,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(247,241,231,0.92)",
    )
    return fig


def build_benchmark_chart(summary_df: pd.DataFrame) -> go.Figure:
    benchmark_df = summary_df.copy()
    benchmark_df["Avg. Duration (min)"] = benchmark_df["Avg. Duration (s)"] / 60
    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=benchmark_df["Sitio"],
            y=benchmark_df["Visits"],
            name="Visits",
            marker_color="#2f241d",
            hovertemplate="%{x}<br>Visits: %{y:,.0f}<extra></extra>",
        )
    )
    fig.add_trace(
        go.Scatter(
            x=benchmark_df["Sitio"],
            y=benchmark_df["Avg. Duration (min)"],
            name="Avg. Duration (min)",
            mode="lines+markers",
            line=dict(color="#c47a2c", width=3),
            marker=dict(size=8),
            yaxis="y2",
            hovertemplate="%{x}<br>Avg. Duration: %{y:.2f} min<extra></extra>",
        )
    )
    fig.update_layout(
        height=420,
        margin=dict(t=24, r=24, b=24, l=24),
        xaxis_title="Sitio",
        yaxis=dict(title="Visits"),
        yaxis2=dict(title="Avg. Duration (min)", overlaying="y", side="right", showgrid=False),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(247,241,231,0.92)",
    )
    return fig


def build_device_mix_chart(summary_df: pd.DataFrame) -> go.Figure:
    device_df = summary_df.sort_values("Mobile (%)", ascending=False).copy()
    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            y=device_df["Sitio"],
            x=device_df["Desktop (%)"],
            name="Desktop",
            orientation="h",
            marker_color="#355f8c",
            text=[f"{value:.1f}%" for value in device_df["Desktop (%)"]],
            textposition="inside",
            insidetextanchor="middle",
            hovertemplate="%{y}<br>Desktop: %{x:.1f}%<extra></extra>",
        )
    )
    fig.add_trace(
        go.Bar(
            y=device_df["Sitio"],
            x=device_df["Mobile (%)"],
            name="Mobile",
            orientation="h",
            marker_color="#1f6f43",
            text=[f"{value:.1f}%" for value in device_df["Mobile (%)"]],
            textposition="inside",
            insidetextanchor="middle",
            hovertemplate="%{y}<br>Mobile: %{x:.1f}%<extra></extra>",
        )
    )
    fig.update_layout(
        height=420,
        margin=dict(t=24, r=24, b=24, l=24),
        barmode="stack",
        xaxis_title="Participacion (%)",
        yaxis_title="Sitio",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(247,241,231,0.92)",
    )
    fig.update_xaxes(range=[0, 100])
    return fig


def build_momentum_chart(filtered_frames: dict[str, pd.DataFrame], focus_competitor: str) -> go.Figure:
    fig = go.Figure()
    for site in [PRIMARY_SITE, focus_competitor]:
        df = filtered_frames[site].sort_values("Date").copy()
        df["WoW Change"] = df["Visits"].pct_change() * 100
        fig.add_trace(
            go.Bar(
                x=df["Date"],
                y=df["WoW Change"],
                name=display_site_name(site),
                marker_color=PALETTE.get(site, "#1a1714"),
                opacity=0.78 if site == PRIMARY_SITE else 0.55,
                hovertemplate="%{x|%d %b %Y}<br>%{y:.1f}% WoW<extra></extra>",
            )
        )
    fig.update_layout(
        height=420,
        margin=dict(t=24, r=24, b=24, l=24),
        barmode="group",
        xaxis_title="Semana",
        yaxis_title="Cambio semanal (%)",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(247,241,231,0.92)",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    fig.add_hline(y=0, line_width=1, line_dash="dash", line_color="rgba(26,23,20,0.35)")
    return fig


def build_scorecard_table(summary_df: pd.DataFrame) -> pd.DataFrame:
    table_df = summary_df.copy()
    table_df["Avg. Duration"] = table_df["Avg. Duration (s)"].apply(format_seconds_label)
    table_df["Visits"] = table_df["Visits"].apply(lambda value: f"{value:,.2f}")
    table_df["Page Views"] = table_df["Page Views"].apply(lambda value: f"{value:,.2f}")
    table_df["Pages / Visit"] = table_df["Pages / Visit"].apply(lambda value: f"{value:.2f}")
    table_df["Bounce Rate (%)"] = table_df["Bounce Rate (%)"].apply(lambda value: f"{value:.2f}%")
    table_df["Desktop (%)"] = table_df["Desktop (%)"].apply(lambda value: f"{value:.2f}%")
    table_df["Mobile (%)"] = table_df["Mobile (%)"].apply(lambda value: f"{value:.2f}%")
    table_df["WoW Latest"] = table_df["WoW Latest (%)"].apply(format_delta)
    return table_df[
        [
            "Sitio",
            "Visits",
            "Page Views",
            "Pages / Visit",
            "Bounce Rate (%)",
            "Avg. Duration",
            "Desktop (%)",
            "Mobile (%)",
            "WoW Latest",
        ]
    ].rename(
        columns={
            "Bounce Rate (%)": "Bounce Rate",
            "Desktop (%)": "Desktop",
            "Mobile (%)": "Mobile",
        }
    )


def build_two_site_cards(summary_df: pd.DataFrame, device_df: pd.DataFrame) -> list[dict[str, str]]:
    rows = summary_df.sort_values("total_visits", ascending=False).reset_index(drop=True)
    leader = rows.iloc[0]
    challenger = rows.iloc[1]
    duration_leader = summary_df.sort_values("visit_duration_seconds", ascending=False).iloc[0]
    mobile_df = summary_df.merge(
        device_df.loc[device_df["device"] == "mobile", ["domain", "share"]],
        on="domain",
        how="left",
    )
    mobile_leader = mobile_df.sort_values("share", ascending=False).iloc[0]
    return [
        {
            "title": "Mayor volumen",
            "headline": f"{leader['domain_label']} | {leader['total_visits']:,.0f} visits",
            "description": f"Supera a {challenger['domain_label']} en alcance acumulado durante marzo-mayo 2026.",
        },
        {
            "title": "Mejor retencion",
            "headline": f"{duration_leader['domain_label']} | {format_seconds_label(float(duration_leader['visit_duration_seconds']))}",
            "description": "Es el dominio con mayor permanencia promedio por visita en el comparativo directo.",
        },
        {
            "title": "Mayor peso mobile",
            "headline": f"{mobile_leader['domain_label']} | {mobile_leader['share'] * 100:.1f}% mobile",
            "description": "Concentra la mezcla de trafico mas orientada a telefono dentro del par comparado.",
        },
    ]


def build_two_site_monthly_chart(monthly_df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    for domain in monthly_df["domain"].unique():
        site_df = monthly_df.loc[monthly_df["domain"] == domain].sort_values("month")
        fig.add_trace(
            go.Scatter(
                x=site_df["month"],
                y=site_df["visits"],
                mode="lines+markers",
                name=display_site_name(domain),
                line=dict(color=PALETTE.get(domain, "#1a1714"), width=3),
                marker=dict(size=8),
                hovertemplate="%{x|%b %Y}<br>%{y:,.0f} visits<extra></extra>",
            )
        )
    fig.update_layout(
        height=420,
        margin=dict(t=24, r=24, b=24, l=24),
        xaxis_title="Mes",
        yaxis_title="Visits",
        hovermode="x unified",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(247,241,231,0.92)",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    return fig


def build_two_site_device_chart(device_df: pd.DataFrame) -> go.Figure:
    pivot_df = (
        device_df.pivot(index="domain_label", columns="device", values="share")
        .fillna(0)
        .reset_index()
    )
    pivot_df["desktop"] = pivot_df["desktop"] * 100
    pivot_df["mobile"] = pivot_df["mobile"] * 100
    pivot_df = pivot_df.sort_values("mobile", ascending=False)

    fig = go.Figure()
    for device, color in [("desktop", "#355f8c"), ("mobile", "#1f6f43")]:
        fig.add_trace(
            go.Bar(
                y=pivot_df["domain_label"],
                x=pivot_df[device],
                name=device.capitalize(),
                orientation="h",
                marker_color=color,
                text=[f"{value:.1f}%" for value in pivot_df[device]],
                textposition="inside",
                insidetextanchor="middle",
                hovertemplate="%{y}<br>" + device.capitalize() + ": %{x:.1f}%<extra></extra>",
            )
        )
    fig.update_layout(
        height=420,
        margin=dict(t=24, r=24, b=24, l=24),
        barmode="stack",
        xaxis_title="Participacion (%)",
        yaxis_title="Dominio",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(247,241,231,0.92)",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    fig.update_xaxes(range=[0, 100])
    return fig


def build_two_site_metric_chart(summary_df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=summary_df["domain_label"],
            y=summary_df["monthly_unique_visitors"],
            name="Monthly unique visitors",
            marker_color="#c47a2c",
            hovertemplate="%{x}<br>%{y:,.0f}<extra></extra>",
        )
    )
    fig.add_trace(
        go.Bar(
            x=summary_df["domain_label"],
            y=summary_df["page_views"],
            name="Page views",
            marker_color="#8a4b18",
            hovertemplate="%{x}<br>%{y:,.0f}<extra></extra>",
        )
    )
    fig.add_trace(
        go.Scatter(
            x=summary_df["domain_label"],
            y=summary_df["pages_per_visit"],
            name="Pages / Visit",
            mode="lines+markers",
            line=dict(color="#1f6f43", width=3),
            marker=dict(size=8),
            yaxis="y2",
            hovertemplate="%{x}<br>%{y:.2f} pages/visit<extra></extra>",
        )
    )
    fig.update_layout(
        height=420,
        margin=dict(t=24, r=24, b=24, l=24),
        xaxis_title="Dominio",
        yaxis=dict(title="Volumen"),
        yaxis2=dict(title="Pages / Visit", overlaying="y", side="right", showgrid=False),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(247,241,231,0.92)",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    return fig


def build_two_site_scorecard(summary_df: pd.DataFrame, device_df: pd.DataFrame) -> pd.DataFrame:
    mobile_map = device_df.loc[device_df["device"] == "mobile", ["domain", "share"]].rename(
        columns={"share": "mobile_share"}
    )
    desktop_map = device_df.loc[device_df["device"] == "desktop", ["domain", "share"]].rename(
        columns={"share": "desktop_share"}
    )
    scorecard = summary_df.merge(mobile_map, on="domain").merge(desktop_map, on="domain")
    scorecard["visit_duration"] = scorecard["visit_duration_seconds"].apply(format_seconds_label)
    scorecard["bounce_rate"] = scorecard["bounce_rate"] * 100
    scorecard["mobile_share"] = scorecard["mobile_share"] * 100
    scorecard["desktop_share"] = scorecard["desktop_share"] * 100
    return scorecard[
        [
            "domain_label",
            "total_visits",
            "monthly_visits",
            "monthly_unique_visitors",
            "visit_duration",
            "pages_per_visit",
            "bounce_rate",
            "desktop_share",
            "mobile_share",
            "page_views",
        ]
    ].rename(
        columns={
            "domain_label": "Sitio",
            "total_visits": "Visits",
            "monthly_visits": "Monthly Visits",
            "monthly_unique_visitors": "Monthly UVs",
            "bounce_rate": "Bounce Rate",
            "desktop_share": "Desktop",
            "mobile_share": "Mobile",
            "page_views": "Page Views",
            "visit_duration": "Avg. Duration",
        }
    )


def build_two_site_insight(summary_df: pd.DataFrame) -> str:
    ordered = summary_df.sort_values("total_visits", ascending=False).reset_index(drop=True)
    leader = ordered.iloc[0]
    second = ordered.iloc[1]
    leader_gap = ((leader["total_visits"] / second["total_visits"]) - 1) * 100 if second["total_visits"] else 0.0
    duration_leader = summary_df.sort_values("visit_duration_seconds", ascending=False).iloc[0]
    bounce_leader = summary_df.sort_values("bounce_rate", ascending=True).iloc[0]
    return (
        f"{leader['domain_label']} lidera en volumen con una ventaja de {leader_gap:.1f}% sobre {second['domain_label']}. "
        f"En retencion, el mejor promedio de permanencia lo tiene {duration_leader['domain_label']}, mientras que el rebote mas bajo corresponde a {bounce_leader['domain_label']}. "
        f"Esta pestaña ayuda a separar si la mejor opcion de marca digital gana por escala, por calidad de sesion o por ambas."
    )


def build_progressive_funnel_values(site_row: pd.Series) -> list[tuple[str, float]]:
    visits = float(site_row["monthly_visits"])
    unique_visitors = float(site_row["monthly_unique_visitors"])
    non_bounced = visits * (1 - float(site_row["bounce_rate"]))
    depth_factor = max(0.0, min(1.0, (float(site_row["pages_per_visit"]) - 1.0) / 2.0))
    deep_engagement = non_bounced * depth_factor
    return [
        ("Visits", visits),
        ("Unique Visitors", unique_visitors),
        ("Non-Bounced", non_bounced),
        ("Deep Engagement", deep_engagement),
    ]


def build_progressive_funnel_chart(site_row: pd.Series, color: str) -> go.Figure:
    funnel_values = build_progressive_funnel_values(site_row)
    stages = [stage for stage, _ in funnel_values]
    values = [value for _, value in funnel_values]
    base = values[0] if values else 1
    text = [f"{stage}<br>{(value / base * 100):.1f}%" for stage, value in funnel_values]

    fig = go.Figure(
        go.Funnel(
            y=stages,
            x=values,
            text=text,
            textposition="inside",
            textinfo="text",
            opacity=0.9,
            marker=dict(color=color, line=dict(color="rgba(255,255,255,0.7)", width=1)),
            connector=dict(line=dict(color="rgba(103,91,78,0.18)", width=1)),
            hovertemplate="%{y}<br>%{x:,.0f}<extra></extra>",
        )
    )
    fig.update_layout(
        height=430,
        margin=dict(t=20, r=18, b=20, l=18),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(247,241,231,0.92)",
        showlegend=False,
    )
    fig.update_yaxes(autorange="reversed")
    return fig


def build_funnel_insight(primary_row: pd.Series, competitor_row: pd.Series) -> str:
    primary_funnel = dict(build_progressive_funnel_values(primary_row))
    competitor_funnel = dict(build_progressive_funnel_values(competitor_row))
    primary_uv = (primary_funnel["Unique Visitors"] / primary_funnel["Visits"] * 100) if primary_funnel["Visits"] else 0.0
    competitor_uv = (
        competitor_funnel["Unique Visitors"] / competitor_funnel["Visits"] * 100
        if competitor_funnel["Visits"]
        else 0.0
    )
    primary_deep = (primary_funnel["Deep Engagement"] / primary_funnel["Visits"] * 100) if primary_funnel["Visits"] else 0.0
    competitor_deep = (
        competitor_funnel["Deep Engagement"] / competitor_funnel["Visits"] * 100
        if competitor_funnel["Visits"]
        else 0.0
    )
    better_uv = primary_row["domain_label"] if primary_uv >= competitor_uv else competitor_row["domain_label"]
    better_deep = primary_row["domain_label"] if primary_deep >= competitor_deep else competitor_row["domain_label"]
    return (
        f"En captacion de audiencia unica, {better_uv} muestra mejor conversion relativa de visits a usuarios unicos. "
        f"En profundidad, {better_deep} retiene una mayor proporcion estimada de sesiones de alta implicacion. "
        f"El stage de Deep Engagement es una aproximacion ejecutiva basada en sesiones no rebotadas y profundidad promedio de paginas por visita."
    )


def render_executive_cards(cards: list[dict[str, str]]) -> None:
    if not cards:
        return
    columns = st.columns(len(cards), vertical_alignment="center")
    for column, card in zip(columns, cards):
        with column:
            st.markdown(
                f"""
                <div class="metric-card">
                    <div class="metric-kicker">{card['title']}</div>
                    <div class="metric-headline">{card['headline']}</div>
                    <div class="metric-copy">{card['description']}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


def render_insight_box(title: str, body: str) -> None:
    st.markdown(
        f"""
        <div style="margin-top:12px;border:1px solid rgba(196,122,44,0.16);border-radius:18px;padding:14px 16px;background:linear-gradient(180deg,#ffffff 0%,#f8f1e7 100%);">
            <div style="font-size:0.8rem;font-weight:700;letter-spacing:0.04em;text-transform:uppercase;color:#c47a2c;margin-bottom:8px;">{title}</div>
            <div style="font-size:0.95rem;line-height:1.55;color:#2f241d;">{body}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def get_target_summary_row(summary_df: pd.DataFrame) -> Optional[pd.Series]:
    rows = summary_df.loc[summary_df["Site"] == PRIMARY_SITE]
    if rows.empty:
        return None
    return rows.iloc[0]


def build_executive_cards(summary_df: pd.DataFrame) -> list[dict[str, str]]:
    if summary_df.empty:
        return []

    reach_row = summary_df.sort_values("Visits", ascending=False).iloc[0]
    engagement_score = (
        summary_df["Pages / Visit"] * 18
        + summary_df["Avg. Duration (s)"] / 12
        - summary_df["Bounce Rate (%)"]
    )
    engagement_row = summary_df.assign(engagement_score=engagement_score).sort_values(
        "engagement_score", ascending=False
    ).iloc[0]
    mobile_row = summary_df.sort_values("Mobile (%)", ascending=False).iloc[0]
    target_row = get_target_summary_row(summary_df)
    target_rank = "N/A"
    if target_row is not None:
        rank_position = int(summary_df["Visits"].rank(ascending=False, method="min")[summary_df["Site"] == PRIMARY_SITE].iloc[0])
        target_rank = f"#{rank_position} de {len(summary_df)}"

    return [
        {
            "title": "Mayor alcance",
            "headline": f"{reach_row['Sitio']} | {reach_row['Visits']:,.0f} visits",
            "description": "Es la marca con mayor volumen acumulado de visitas dentro del benchmark activo.",
        },
        {
            "title": "Mejor engagement",
            "headline": f"{engagement_row['Sitio']} | {engagement_row['Pages / Visit']:.2f} pags/visita",
            "description": "Combina profundidad de navegacion, duracion y rebote relativo en una sola lectura ejecutiva.",
        },
        {
            "title": "Mayor peso mobile",
            "headline": f"{mobile_row['Sitio']} | {mobile_row['Mobile (%)']:.1f}% mobile",
            "description": "Es el dominio mas dependiente de sesiones en telefono dentro de la ventana filtrada.",
        },
        {
            "title": "Posicion Loco",
            "headline": f"Loco Tequila | {target_rank}",
            "description": "Ranking de locotequila.mx por volumen de visits frente a su set competitivo.",
        },
    ]


def build_visits_trend_insight(summary_df: pd.DataFrame, filtered_visits_df: pd.DataFrame) -> str:
    target_row = get_target_summary_row(summary_df)
    if target_row is None or filtered_visits_df.empty:
        return "No hay informacion suficiente para leer la evolucion semanal de Loco Tequila frente al benchmark."
    rank_df = summary_df.sort_values("Visits", ascending=False).reset_index(drop=True)
    target_rank = int(rank_df.index[rank_df["Site"] == PRIMARY_SITE][0]) + 1
    leader = rank_df.iloc[0]
    target_series = filtered_visits_df[PRIMARY_SITE]
    peak_date = target_series.idxmax()
    peak_value = float(target_series.max())
    recent = float(target_series.iloc[-1]) if len(target_series) else 0.0
    leader_gap_pct = ((float(target_row["Visits"]) / float(leader["Visits"])) * 100) if float(leader["Visits"]) else 0.0
    return (
        f"Loco Tequila ocupa la posicion #{target_rank} por volumen acumulado de visits en el periodo filtrado. "
        f"Su escala equivale a {leader_gap_pct:.1f}% del lider actual ({leader['Sitio']}). "
        f"Su mejor semana fue la del {peak_date:%d %b %Y} con {peak_value:,.0f} visits, mientras que su ultimo dato visible marca {recent:,.0f} visits. "
        f"Esta lectura ayuda a distinguir si la marca esta construyendo traccion sostenida o si depende de repuntes aislados."
    )


def build_share_insight(summary_df: pd.DataFrame) -> str:
    target_row = get_target_summary_row(summary_df)
    if target_row is None or summary_df.empty:
        return "No hay suficiente informacion para leer el share relativo de Loco Tequila."
    total_visits = float(summary_df["Visits"].sum())
    share = (float(target_row["Visits"]) / total_visits * 100) if total_visits else 0.0
    return (
        f"Loco Tequila captura solo {share:.2f}% del trafico total del benchmark, lo que evidencia una brecha estructural de visibilidad digital frente al resto del set competitivo. "
        f"Estratégicamente, esto sugiere que la marca todavía no está convirtiendo su propuesta premium en alcance consistente a escala, por lo que la prioridad no debería centrarse solo en optimizar engagement, sino en fortalecer la parte alta del embudo: awareness, descubrimiento y adquisición calificada."
    )


def build_competitor_insight(filtered_frames: dict[str, pd.DataFrame], competitor_site: str) -> str:
    if PRIMARY_SITE not in filtered_frames or competitor_site not in filtered_frames:
        return "Activa Loco Tequila y un benchmark directo para obtener un comparativo puntual."
    target_visits = float(filtered_frames[PRIMARY_SITE]["Visits"].sum())
    comp_visits = float(filtered_frames[competitor_site]["Visits"].sum())
    ratio = (target_visits / comp_visits * 100) if comp_visits else 0.0
    diff = target_visits - comp_visits
    direction = "por debajo de" if diff < 0 else "por encima de"
    return (
        f"Frente a {display_site_name(competitor_site)}, Loco Tequila se ubica {direction} su volumen acumulado y hoy representa {ratio:.1f}% de esa escala. "
        f"La utilidad estratégica de esta vista está en separar dos escenarios: una distancia estable que exige construir awareness, o una brecha más volátil que abre oportunidades tácticas de captura."
    )


def build_heatmap_insight(summary_df: pd.DataFrame) -> str:
    target_row = get_target_summary_row(summary_df)
    if target_row is None or len(summary_df) <= 1:
        return "No hay suficientes sitios para comparar el engagement de Loco Tequila."
    avg_pages = float(summary_df["Pages / Visit"].mean())
    avg_bounce = float(summary_df["Bounce Rate (%)"].mean())
    avg_duration = float(summary_df["Avg. Duration (s)"].mean())
    return (
        f"Loco Tequila registra {target_row['Pages / Visit']:.2f} paginas por visita frente a un promedio benchmark de {avg_pages:.2f}. "
        f"Su bounce rate es {target_row['Bounce Rate (%)']:.1f}% contra {avg_bounce:.1f}% del conjunto, y su duracion media es {format_seconds_label(float(target_row['Avg. Duration (s)']))} frente a {format_seconds_label(avg_duration)} del benchmark. "
        f"El heatmap deja ver si la marca compensa menor escala con sesiones mas calificadas."
    )


def build_quality_scatter_insight(summary_df: pd.DataFrame) -> str:
    target_row = get_target_summary_row(summary_df)
    if target_row is None:
        return "No hay suficientes datos para ubicar a Loco Tequila en la matriz de calidad."
    avg_pages = float(summary_df["Pages / Visit"].mean())
    avg_bounce = float(summary_df["Bounce Rate (%)"].mean())
    pages_position = "por encima" if float(target_row["Pages / Visit"]) >= avg_pages else "por debajo"
    bounce_position = "mejor" if float(target_row["Bounce Rate (%)"]) <= avg_bounce else "peor"
    return (
        f"Loco Tequila aparece con {target_row['Pages / Visit']:.2f} paginas por visita y {target_row['Bounce Rate (%)']:.1f}% de rebote. "
        f"Eso la coloca {pages_position} del promedio en profundidad y con un desempeno {bounce_position} que el benchmark en rebote. "
        f"Si su punto se mueve hacia arriba y a la izquierda, la marca estaria combinando mejor calidad de trafico con sesiones mas profundas."
    )


def build_benchmark_insight(summary_df: pd.DataFrame) -> str:
    target_row = get_target_summary_row(summary_df)
    if target_row is None or summary_df.empty:
        return "No hay suficiente informacion para contrastar alcance y permanencia de Loco Tequila."
    visits_rank = int(summary_df["Visits"].rank(ascending=False, method="min")[summary_df["Site"] == PRIMARY_SITE].iloc[0])
    duration_rank = int(summary_df["Avg. Duration (s)"].rank(ascending=False, method="min")[summary_df["Site"] == PRIMARY_SITE].iloc[0])
    return (
        f"Loco Tequila ocupa la posicion #{visits_rank} en alcance y la posicion #{duration_rank} en permanencia promedio. "
        f"Esta relacion separa dos preguntas ejecutivas: cuanto trafico logra atraer la marca y que tan bien retiene ese trafico una vez dentro."
    )


def build_device_mix_insight(summary_df: pd.DataFrame) -> str:
    target_row = get_target_summary_row(summary_df)
    if target_row is None:
        return "No hay suficiente informacion para leer la mezcla de dispositivo."
    mobile_rank = int(summary_df["Mobile (%)"].rank(ascending=False, method="min")[summary_df["Site"] == PRIMARY_SITE].iloc[0])
    return (
        f"Loco Tequila promedia {target_row['Mobile (%)']:.1f}% de trafico mobile y {target_row['Desktop (%)']:.1f}% desktop, ubicandose en la posicion #{mobile_rank} del benchmark por peso movil. "
        f"Esto ayuda a decidir si la experiencia principal a optimizar es la navegacion en telefono o si todavia existe una base desktop relativamente relevante."
    )


def build_momentum_insight(filtered_frames: dict[str, pd.DataFrame], competitor_site: str) -> str:
    if PRIMARY_SITE not in filtered_frames or competitor_site not in filtered_frames:
        return "No hay suficiente informacion para leer el momentum semanal."
    target_df = filtered_frames[PRIMARY_SITE].sort_values("Date").copy()
    competitor_df = filtered_frames[competitor_site].sort_values("Date").copy()
    target_df["wow"] = target_df["Visits"].pct_change() * 100
    competitor_df["wow"] = competitor_df["Visits"].pct_change() * 100
    target_peak = target_df.loc[target_df["wow"].idxmax()] if target_df["wow"].notna().any() else None
    comp_peak = competitor_df.loc[competitor_df["wow"].idxmax()] if competitor_df["wow"].notna().any() else None
    if target_peak is None or comp_peak is None:
        return "Todavia no hay suficientes puntos semanales para leer aceleraciones o retrocesos."
    return (
        f"El mayor salto semanal de Loco Tequila fue {target_peak['wow']:.1f}% en la semana del {target_peak['Date']:%d %b %Y}. "
        f"En {display_site_name(competitor_site)}, el mejor repunte fue {comp_peak['wow']:.1f}% en {comp_peak['Date']:%d %b %Y}. "
        f"Esta vista permite detectar que marca trae mejor impulso reciente, mas alla del volumen absoluto."
    )


def render_filters(
    combined_df: pd.DataFrame,
    source_frames: dict[str, pd.DataFrame],
) -> tuple[pd.Timestamp, pd.Timestamp, list[str], bool]:
    min_date = combined_df.index.min().date()
    max_date = combined_df.index.max().date()
    has_partial_last_week = pd.Timestamp(max_date).weekday() != 6

    st.markdown('<div class="filters-shell">', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1.15, 1.3, 0.75])
    with col1:
        selected_dates = st.date_input(
            "Rango de fechas",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
        )
    with col2:
        selected_sites = st.multiselect(
            "Sitios",
            options=list(source_frames.keys()),
            default=list(source_frames.keys()),
            format_func=display_site_name,
        )
    with col3:
        if has_partial_last_week:
            exclude_partial_week = st.checkbox("Excluir semana parcial final", value=True)
        else:
            exclude_partial_week = False
    st.markdown("</div>", unsafe_allow_html=True)

    if isinstance(selected_dates, tuple) and len(selected_dates) == 2:
        start_date, end_date = selected_dates
    else:
        start_date = end_date = selected_dates

    if PRIMARY_SITE not in selected_sites:
        selected_sites = [PRIMARY_SITE] + selected_sites
    if not selected_sites:
        selected_sites = list(source_frames.keys())

    return pd.Timestamp(start_date), pd.Timestamp(end_date), selected_sites, exclude_partial_week


def filter_source_frames(
    source_frames: dict[str, pd.DataFrame],
    selected_sites: list[str],
    start_date: pd.Timestamp,
    end_date: pd.Timestamp,
    exclude_partial_week: bool,
) -> dict[str, pd.DataFrame]:
    filtered: dict[str, pd.DataFrame] = {}
    for site, frame in source_frames.items():
        if site not in selected_sites:
            continue
        df = frame.loc[(frame["Date"] >= start_date) & (frame["Date"] <= end_date)].copy()
        if exclude_partial_week and len(df) > 1:
            max_date = df["Date"].max()
            if max_date.weekday() != 6:
                df = df.loc[df["Date"] != max_date].copy()
        filtered[site] = df
    return filtered


def render_benchmark_dashboard() -> None:
    source_frames = load_source_frames()
    combined_visits_df = load_combined_visits(source_frames)
    competitors = [site for site in source_frames if site != PRIMARY_SITE]

    start_date, end_date, selected_sites, exclude_partial_week = render_filters(
        combined_visits_df,
        source_frames,
    )

    filtered_frames = filter_source_frames(
        source_frames,
        selected_sites,
        start_date,
        end_date,
        exclude_partial_week,
    )
    filtered_visits_df = combined_visits_df.loc[start_date:end_date, selected_sites].copy()
    if exclude_partial_week and len(filtered_visits_df) > 1:
        max_date = filtered_visits_df.index.max()
        if max_date.weekday() != 6:
            filtered_visits_df = filtered_visits_df.loc[filtered_visits_df.index != max_date]

    summary_df = build_summary_table(filtered_frames)
    render_executive_cards(build_executive_cards(summary_df))
    st.markdown("<div style='height:18px;'></div>", unsafe_allow_html=True)

    open_chart_shell(
        "Visits por semana",
        "Vista principal para detectar ritmo competitivo, picos de adquisicion y la distancia relativa de Loco Tequila frente a marcas con mucha mayor escala.",
        featured=True,
    )
    if filtered_visits_df.empty:
        st.info("No hay datos de visits disponibles para los filtros seleccionados.")
    else:
        st.plotly_chart(build_visits_trend_chart(filtered_visits_df), use_container_width=True)
        render_insight_box("Insight", build_visits_trend_insight(summary_df, filtered_visits_df))
    close_chart_shell()

    left, right = st.columns([1.1, 0.9], vertical_alignment="top")
    with left:
        open_chart_shell(
            "Share of traffic",
            "Distribucion del volumen de visits entre los dominios activos en el periodo filtrado.",
        )
        if summary_df.empty:
            st.info("No hay datos para calcular share of traffic.")
        else:
            st.plotly_chart(build_traffic_share_chart(summary_df), use_container_width=True)
            render_insight_box("Insight", build_share_insight(summary_df))
        close_chart_shell()
    with right:
        open_chart_shell(
            "Loco vs benchmark directo",
            "Comparacion puntual de visits entre Loco Tequila y el competidor seleccionado.",
        )
        compare_header_left, compare_header_right = st.columns([0.78, 0.22], vertical_alignment="center")
        with compare_header_left:
            st.markdown(
                "<div style='font-size:0.82rem;color:#675b4e;margin:0 0 8px 2px;'>Selecciona el rival que quieres contrastar contra Loco Tequila.</div>",
                unsafe_allow_html=True,
            )
        with compare_header_right:
            focus_competitor = st.selectbox(
                "Benchmark directo",
                options=competitors,
                index=0,
                format_func=display_site_name,
                label_visibility="collapsed",
                key="direct_benchmark_selector",
            )
        if PRIMARY_SITE in filtered_frames and focus_competitor in filtered_frames:
            st.plotly_chart(
                build_target_vs_competitor_chart(filtered_frames, focus_competitor),
                use_container_width=True,
            )
            render_insight_box("Insight", build_competitor_insight(filtered_frames, focus_competitor))
        else:
            st.info("Manten activos a Loco Tequila y al benchmark directo para esta vista.")
        close_chart_shell()

    center_left, center_right = st.columns([1.05, 0.95], vertical_alignment="top")
    with center_left:
        open_chart_shell(
            "Heatmap de engagement",
            "Lectura comparada de profundidad, rebote, duracion y mezcla de dispositivo.",
        )
        if summary_df.empty:
            st.info("No hay KPIs disponibles para mostrar.")
        else:
            st.plotly_chart(build_kpi_heatmap(summary_df), use_container_width=True)
            render_insight_box("Insight", build_heatmap_insight(summary_df))
        close_chart_shell()
    with center_right:
        open_chart_shell(
            "Scatter de calidad de trafico",
            "Cruza rebote, paginas por visita y escala para detectar trafico mas profundo contra trafico mas superficial.",
        )
        if summary_df.empty:
            st.info("No hay datos suficientes para esta comparacion.")
        else:
            st.plotly_chart(build_quality_scatter(summary_df), use_container_width=True)
            render_insight_box("Insight", build_quality_scatter_insight(summary_df))
        close_chart_shell()

    open_chart_shell(
        "Reach vs permanencia",
        "Relacion entre el volumen de visits y la duracion promedio para ubicar sitios con escala y mejor retencion.",
    )
    if summary_df.empty:
        st.info("No hay datos suficientes para esta vista.")
    else:
        st.plotly_chart(build_benchmark_chart(summary_df), use_container_width=True)
        render_insight_box("Insight", build_benchmark_insight(summary_df))
    close_chart_shell()

    device_left, device_right = st.columns([1.0, 1.0], vertical_alignment="top")
    with device_left:
        open_chart_shell(
            "Mix por dispositivo",
            "Comparativo del peso desktop y mobile por marca para leer donde vive realmente el consumo.",
        )
        if summary_df.empty:
            st.info("No hay datos suficientes para esta vista.")
        else:
            st.plotly_chart(build_device_mix_chart(summary_df), use_container_width=True)
            render_insight_box("Insight", build_device_mix_insight(summary_df))
        close_chart_shell()
    with device_right:
        open_chart_shell(
            "Momentum semanal",
            "Variacion porcentual semana contra semana para separar escala de impulso reciente.",
        )
        momentum_header_left, momentum_header_right = st.columns([0.78, 0.22], vertical_alignment="center")
        with momentum_header_left:
            st.markdown(
                "<div style='font-size:0.82rem;color:#675b4e;margin:0 0 8px 2px;'>Selecciona un competidor para contrastar el impulso semanal de Loco Tequila sin afectar otras vistas.</div>",
                unsafe_allow_html=True,
            )
        with momentum_header_right:
            momentum_competitor = st.selectbox(
                "Competidor momentum",
                options=competitors,
                index=0,
                format_func=display_site_name,
                label_visibility="collapsed",
                key="momentum_competitor_selector",
            )
        if (
            PRIMARY_SITE in filtered_frames
            and momentum_competitor in filtered_frames
            and len(filtered_frames[PRIMARY_SITE]) > 1
        ):
            st.plotly_chart(
                build_momentum_chart(filtered_frames, momentum_competitor),
                use_container_width=True,
            )
            render_insight_box("Insight", build_momentum_insight(filtered_frames, momentum_competitor))
        else:
            st.info("Se requieren al menos dos semanas visibles para comparar momentum.")
        close_chart_shell()

    open_chart_shell(
        "Funnels progresivos de engagement",
        "Comparativo de conversion relativa desde visits hacia audiencia unica, sesiones no rebotadas y una capa estimada de deep engagement.",
    )
    try:
        competitor_summary = load_competitor_database_summary()
        funnel_competitors = [
            site for site in competitor_summary["domain"].tolist() if site != PRIMARY_SITE
        ]
        funnel_header_left, funnel_header_right = st.columns([0.78, 0.22], vertical_alignment="center")
        with funnel_header_left:
            st.markdown(
                "<div style='font-size:0.82rem;color:#675b4e;margin:0 0 8px 2px;'>Loco Tequila permanece fijo; a la derecha puedes contrastar cualquier competidor sin cambiar el resto del dashboard.</div>",
                unsafe_allow_html=True,
            )
        with funnel_header_right:
            selected_funnel_competitor = st.selectbox(
                "Competidor funnel",
                options=funnel_competitors,
                index=0,
                format_func=display_site_name,
                label_visibility="collapsed",
                key="funnel_competitor_selector",
            )

        primary_funnel_row = competitor_summary.loc[competitor_summary["domain"] == PRIMARY_SITE].iloc[0]
        competitor_funnel_row = competitor_summary.loc[
            competitor_summary["domain"] == selected_funnel_competitor
        ].iloc[0]

        funnel_left, funnel_right = st.columns(2, vertical_alignment="top")
        with funnel_left:
            st.markdown(
                f"<div style='font-size:1rem;font-weight:700;color:#1a1714;margin:0 0 8px 4px;'>{primary_funnel_row['domain_label']}</div>",
                unsafe_allow_html=True,
            )
            st.plotly_chart(
                build_progressive_funnel_chart(primary_funnel_row, PALETTE.get(PRIMARY_SITE, "#c47a2c")),
                use_container_width=True,
            )
        with funnel_right:
            st.markdown(
                f"<div style='font-size:1rem;font-weight:700;color:#1a1714;margin:0 0 8px 4px;'>{competitor_funnel_row['domain_label']}</div>",
                unsafe_allow_html=True,
            )
            st.plotly_chart(
                build_progressive_funnel_chart(
                    competitor_funnel_row,
                    PALETTE.get(selected_funnel_competitor, "#355f8c"),
                ),
                use_container_width=True,
            )

        render_insight_box("Insight", build_funnel_insight(primary_funnel_row, competitor_funnel_row))
        st.caption(
            "Los funnels usan el panel mensual consolidado del benchmark. Deep Engagement es una estimacion ejecutiva basada en sesiones no rebotadas y profundidad promedio de paginas por visita."
        )
    except Exception as exc:
        st.info(f"No fue posible construir los funnels comparativos: {exc}")
    close_chart_shell()

    open_chart_shell(
        "Scorecard comparativo",
        "Resumen tabular con los KPIs clave del benchmark digital del periodo filtrado.",
    )
    if summary_df.empty:
        st.info("No hay scorecard disponible para los filtros seleccionados.")
    else:
        st.dataframe(
            build_scorecard_table(summary_df),
            width="stretch",
            hide_index=True,
        )
    close_chart_shell()

    if exclude_partial_week:
        st.caption(
            "La vista excluye por defecto la ultima semana parcial del archivo para evitar distorsiones por periodos incompletos."
        )


def render_two_site_dashboard() -> None:
    data = load_two_site_data()
    summary_df = data["summary"]
    monthly_df = data["monthly"]
    device_df = data["device"]

    render_executive_cards(build_two_site_cards(summary_df, device_df))
    st.markdown("<div style='height:18px;'></div>", unsafe_allow_html=True)

    open_chart_shell(
        "Loco Tequila .com vs .mx",
        "Comparativo aislado entre ambos dominios para leer diferencias de escala, retencion y mezcla de dispositivo sin ruido competitivo externo.",
        featured=True,
    )
    st.plotly_chart(build_two_site_monthly_chart(monthly_df), use_container_width=True)
    render_insight_box("Insight", build_two_site_insight(summary_df))
    close_chart_shell()

    left, right = st.columns(2, vertical_alignment="top")
    with left:
        open_chart_shell(
            "Mix por dispositivo",
            "Peso relativo de desktop y mobile para cada dominio en el periodo marzo-mayo 2026.",
        )
        st.plotly_chart(build_two_site_device_chart(device_df), use_container_width=True)
        close_chart_shell()
    with right:
        open_chart_shell(
            "Audiencia y profundidad",
            "Cruza visitantes unicos, page views y paginas por visita para leer volumen util contra profundidad.",
        )
        st.plotly_chart(build_two_site_metric_chart(summary_df), use_container_width=True)
        close_chart_shell()

    open_chart_shell(
        "Scorecard .com vs .mx",
        "Resumen tabular para comparar lado a lado las metricas principales del periodo.",
    )
    st.dataframe(build_two_site_scorecard(summary_df, device_df), width="stretch", hide_index=True)
    close_chart_shell()


def render_app() -> None:
    apply_dashboard_styles()
    st.markdown(
        """
        <div class="dashboard-hero">
            <div class="hero-brand">
                <img class="hero-logo" src="https://www.locotequila.mx/hubfs/Loco_Tequila_Logo.svg" alt="Loco Tequila logo">
                <div class="hero-copy">
                    <div class="hero-kicker">Nuestra Colección | Terruño | Proceso</div>
                    <h1>WTA Loco Tequila</h1>
                </div>
            </div>
            <p>Dashboard estratégico para entender la posición digital de Loco Tequila frente a su benchmark competitivo, identificar brechas de alcance y calidad de tráfico, y detectar dónde existe mayor oportunidad para fortalecer visibilidad, consideración y profundidad de engagement.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    benchmark_tab, dual_tab = st.tabs(
        ["Benchmark competitivo", "Loco Tequila .com vs .mx"]
    )
    with benchmark_tab:
        render_benchmark_dashboard()
    with dual_tab:
        render_two_site_dashboard()
