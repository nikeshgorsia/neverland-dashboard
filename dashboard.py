import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(
    page_title="Neverland Finance",
    page_icon="assets/Neverland_Pink_N_1x1.png",
    layout="wide",
)

# ── Neverland VI ──────────────────────────────────────────────────────────────
NV_PINK    = "#E8147B"
NV_DARK    = "#1A1A1A"
NV_LIGHT   = "#F9F9F9"
NV_GREY    = "#F0F0F0"

# Load Futura PT font faces (only available locally, not on cloud)
try:
    with open("assets/futura_fonts.css", "r") as _f:
        _font_css = _f.read()
except FileNotFoundError:
    _font_css = ""

# Load loading background
import base64 as _b64
with open("assets/loading_bg.jpg", "rb") as _f:
    _bg_b64 = _b64.b64encode(_f.read()).decode()
with open("assets/Neverland_Logo_Pink.png", "rb") as _f:
    _logo_b64 = _b64.b64encode(_f.read()).decode()

st.markdown(f"<style>{_font_css}</style>", unsafe_allow_html=True)

st.markdown(f"""
<style>
    /* Loading / splash screen */
    [data-testid="stAppViewContainer"] > .main::before {{
        content: '';
        position: fixed;
        inset: 0;
        background: url('data:image/jpeg;base64,{_bg_b64}') center/cover no-repeat;
        z-index: -1;
        opacity: 1;
    }}

    /* Hide default spinner text */
    .stSpinner p, [data-testid="stSpinner"] p {{
        display: none !important;
    }}

    /* Apply Futura PT font */
    html, body, [class*="css"] {{
        font-family: 'FuturaPT-Book', sans-serif !important;
    }}
    h1, h2, h3, h4, [data-testid="stMetricValue"] {{
        font-family: 'FuturaPT-Bold', sans-serif !important;
    }}
    [data-testid="stMetricValue"] {{
        color: {NV_PINK} !important;
    }}

    /* Hide Stop and Deploy buttons, keep 3-dot menu */
    [data-testid="stToolbarActions"] button:not([aria-label="Open menu"]) {{
        display: none !important;
    }}
    [data-testid="stToolbarActions"] a {{
        display: none !important;
    }}
    [data-testid="stStatusWidget"] {{
        display: none !important;
    }}
    #MainMenu {{
        visibility: hidden;
    }}

    /* Active tab — pink underline */
    [data-testid="stTabs"] [role="tab"][aria-selected="true"] {{
        color: {NV_PINK} !important;
        border-bottom: 3px solid {NV_PINK} !important;
    }}

    /* Metric cards — black background, white text */
    [data-testid="stMetric"] {{
        background: {NV_DARK} !important;
        border-radius: 8px;
        padding: 12px 16px;
        border-left: 4px solid {NV_PINK};
    }}
    [data-testid="stMetricLabel"] p {{
        color: #FFFFFF !important;
        font-weight: 700 !important;
    }}
    [data-testid="stMetricDelta"] {{
        font-weight: 700 !important;
        font-size: 1rem !important;
    }}

    /* Sidebar buttons — pink */
    [data-testid="stSidebar"] .stButton > button {{
        background-color: {NV_PINK} !important;
        color: white !important;
        border: none !important;
        border-radius: 6px !important;
        font-family: 'FuturaPT-Bold', sans-serif !important;
    }}
    [data-testid="stSidebar"] .stButton > button:hover {{
        background-color: #C0106A !important;
    }}
</style>
""", unsafe_allow_html=True)

import time

try:
    from sharepoint_sync import get_device_flow, complete_device_flow, fetch_pipeline, fetch_raw, fetch_client_projects, fetch_budget, fetch_capacity, fetch_scope_debug, fetch_scope, fetch_scope_structure_debug, fetch_salary_by_dept, save_pipeline_snapshot, list_pipeline_snapshots, load_pipeline_snapshot, save_sow_data, load_sow_data
    SHAREPOINT_AVAILABLE = True
except ImportError:
    SHAREPOINT_AVAILABLE = False

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

REFRESH_INTERVAL = 1800  # auto-refresh every 30 minutes

# ── Cross-session data cache (shared across users, 30-min TTL) ────────────────
# The _token arg is excluded from the cache key (leading underscore) so all
# sessions within the TTL window return the same cached result without re-fetching.
if SHAREPOINT_AVAILABLE:
    @st.cache_data(ttl=REFRESH_INTERVAL, show_spinner=False)
    def _cf_pipeline(_token):
        return fetch_pipeline(_token)

    @st.cache_data(ttl=REFRESH_INTERVAL, show_spinner=False)
    def _cf_budget(_token):
        return fetch_budget(_token)

    @st.cache_data(ttl=REFRESH_INTERVAL, show_spinner=False)
    def _cf_capacity(_token):
        return fetch_capacity(_token)

    @st.cache_data(ttl=REFRESH_INTERVAL, show_spinner=False)
    def _cf_scope(_token):
        return fetch_scope(_token)

    @st.cache_data(ttl=REFRESH_INTERVAL, show_spinner=False)
    def _cf_salary(_token):
        return fetch_salary_by_dept(_token)

    def _clear_data_caches():
        _cf_pipeline.clear()
        _cf_budget.clear()
        _cf_capacity.clear()
        _cf_scope.clear()
        _cf_salary.clear()

# ── Restore token + auto-sync on every cold open ──────────────────────────────
if SHAREPOINT_AVAILABLE and "sp_token" not in st.session_state:
    try:
        _, _, cached_token = get_device_flow()
        if cached_token:
            st.session_state["sp_token"] = cached_token
            st.session_state.pop("last_refresh", None)  # force immediate sync
    except Exception:
        pass

# ── Auto-refresh: two-phase loading ──────────────────────────────────────────
# Phase 1 (fast): load pipeline only → show dashboard immediately.
# Phase 2 (deferred): load remaining sources in the background after phase 1.
# After the first load, st.cache_data keeps results for 30 min across all
# sessions/users — subsequent loads return from cache in milliseconds.
_LOADING_HTML = (
    "<style>"
    "@keyframes clouds-drift{0%{transform:scale(1.1) translateX(0px)}100%{transform:scale(1.1) translateX(-120px)}}"
    "@keyframes sweep{0%{clip-path:inset(0 100% 0 0)}60%{clip-path:inset(0 0% 0 0)}85%{clip-path:inset(0 0% 0 0);opacity:1}100%{clip-path:inset(0 0% 0 0);opacity:0}}"
    f".nv-bg{{position:fixed;inset:-10%;background:url('data:image/jpeg;base64,{_bg_b64}') center/cover no-repeat;animation:clouds-drift 20s linear infinite;z-index:0}}"
    ".nv-lc{position:relative;z-index:1;display:flex;flex-direction:column;align-items:center;gap:10px}"
    ".nv-lt{font-family:'FuturaPT-Bold',sans-serif;font-size:1.4rem;color:#FFF;letter-spacing:.18em;text-transform:uppercase;animation:sweep 2s ease-in-out infinite;white-space:nowrap;text-shadow:0 2px 12px rgba(0,0,0,.25)}"
    "</style>"
    "<div style='position:fixed;inset:0;z-index:9999;overflow:hidden;display:flex;align-items:center;justify-content:center'>"
    "<div class='nv-bg'></div>"
    "<div class='nv-lc'>"
    f"<img src='data:image/png;base64,{_logo_b64}' style='width:300px;filter:drop-shadow(0 4px 24px rgba(0,0,0,.2))'>"
    "<p class='nv-lt' style='margin-top:8px'>Loading your dashboard...</p>"
    "</div></div>"
)

if SHAREPOINT_AVAILABLE and "sp_token" in st.session_state:
    from concurrent.futures import ThreadPoolExecutor, as_completed
    last_refresh  = st.session_state.get("last_refresh", 0)
    now           = time.time()
    pipeline_stale = ("pipeline" not in st.session_state) or (now - last_refresh > REFRESH_INTERVAL)
    secondary_needed = "pipeline" in st.session_state and (
        "budget" not in st.session_state or
        "capacity" not in st.session_state or
        "scope" not in st.session_state or
        "salary_dept" not in st.session_state
    )

    if (pipeline_stale or secondary_needed) and not st.session_state.get("_refreshing"):
        st.session_state["_refreshing"] = True
        token = st.session_state["sp_token"]

        if pipeline_stale:
            # Phase 1: pipeline only — show loading screen, then render dashboard fast
            st.markdown(_LOADING_HTML, unsafe_allow_html=True)
            with st.spinner(""):
                try:
                    st.session_state["pipeline"] = _cf_pipeline(token)
                    st.session_state["last_refresh"] = now
                    try:
                        save_pipeline_snapshot(st.session_state["pipeline"])
                    except Exception:
                        pass
                except Exception as e:
                    st.warning(f"Pipeline load failed: {e}")
            st.session_state["_refreshing"] = False
            st.rerun()

        elif secondary_needed:
            # Phase 2: silently load remaining sources (no full-screen overlay)
            with st.spinner("Loading supporting data..."):
                try:
                    tasks = {
                        "budget":      lambda: _cf_budget(token),
                        "capacity":    lambda: _cf_capacity(token),
                        "scope":       lambda: _cf_scope(token),
                        "salary_dept": lambda: _cf_salary(token),
                    }
                    with ThreadPoolExecutor(max_workers=4) as executor:
                        futures = {executor.submit(fn): key for key, fn in tasks.items()}
                        for future in as_completed(futures):
                            key = futures[future]
                            try:
                                st.session_state[key] = future.result()
                            except Exception as e:
                                st.warning(f"{key} load failed: {e}")
                except Exception:
                    pass
            st.session_state["_refreshing"] = False
            st.rerun()

SECTION_LABELS = {
    "CONFIRMED":       "Confirmed Revenue",
    "PROPOSED":        "Proposed",
    "POTENTIAL":       "Account Planning",
    "ACCOUNT PLANNING":"Account Planning",
    "SPECULATIVE":     "Speculative",
}

SECTION_COLORS = {
    "CONFIRMED":        px.colors.sequential.Greens_r,
    "PROPOSED":         px.colors.sequential.Blues_r,
    "POTENTIAL":        px.colors.sequential.Purples_r,
    "ACCOUNT PLANNING": px.colors.sequential.Purples_r,
    "SPECULATIVE":      px.colors.sequential.Oranges_r,
}

def fmt_gbp(v):
    if abs(v) >= 1_000_000:
        return f"£{v/1_000_000:.2f}M"
    if abs(v) >= 1_000:
        return f"£{v/1_000:.0f}k"
    return f"£{v:,.0f}"


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.image("assets/Neverland_Logo_Black.png", use_container_width=True)
    st.divider()

    if SHAREPOINT_AVAILABLE:
        if st.button("🔄 Sync All Data", use_container_width=True):
            try:
                app, flow, token = get_device_flow()
                if token:
                    st.session_state["sp_token"] = token
                    st.session_state.pop("sp_flow", None)
                    if "_token_cache_str" in st.session_state:
                        st.session_state["_show_token_setup"] = True
                    # Clear data caches so fresh data is fetched
                    _clear_data_caches()
                    for _k in ("pipeline", "budget", "capacity", "scope", "salary_dept"):
                        st.session_state.pop(_k, None)
                    st.session_state.pop("last_refresh", None)
                    st.rerun()
                else:
                    st.session_state["sp_flow"] = (app, flow)
            except Exception as e:
                st.error(f"Error: {e}")

        if "sp_flow" in st.session_state:
            app, flow = st.session_state["sp_flow"]
            st.warning("**Sign in required**")
            st.markdown("1. Go to **[microsoft.com/devicelogin](https://microsoft.com/devicelogin)**")
            st.code(flow["user_code"], language=None)
            st.caption("Paste the code above, sign in, then click below.")
            if st.button("✅ I've signed in — load my data", use_container_width=True):
                with st.spinner("Completing login..."):
                    try:
                        token = complete_device_flow(app, flow)
                        st.session_state["sp_token"] = token
                        st.session_state.pop("sp_flow", None)
                        st.session_state.pop("last_refresh", None)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Login failed: {e}")


    if "last_refresh" in st.session_state:
        import datetime
        ts = datetime.datetime.fromtimestamp(st.session_state["last_refresh"]).strftime("%H:%M:%S")
        st.caption(f"Last updated: {ts}")


    # One-time token setup helper
    if st.session_state.get("_show_token_setup") and "_token_cache_str" in st.session_state:
        with st.expander("⚙️ Save login for future sessions", expanded=True):
            st.caption("To avoid signing in every time, add this to your Streamlit secrets:")
            st.code(f'TOKEN_CACHE = {repr(st.session_state["_token_cache_str"])}', language="toml")
            st.caption("Go to share.streamlit.io → your app → ⋮ → Settings → Secrets → paste above → Save")
            if st.button("✅ Done, dismiss", key="dismiss_token"):
                st.session_state.pop("_show_token_setup", None)
                st.rerun()

    st.divider()

    VIEW_OPTIONS = {
        "Confirmed":                               ["CONFIRMED"],
        "Confirmed + Proposed":                    ["CONFIRMED", "PROPOSED"],
        "Confirmed + Proposed + Account Planning": ["CONFIRMED", "PROPOSED", "POTENTIAL", "ACCOUNT PLANNING"],
        "ALL":                                     ["CONFIRMED", "PROPOSED", "POTENTIAL", "ACCOUNT PLANNING", "SPECULATIVE"],
    }
    PL_VIEWS = VIEW_OPTIONS  # same dict used by P&L and BVA tabs

    if "pipeline" in st.session_state:
        pipeline_keys = list(st.session_state["pipeline"].keys())
        if pipeline_keys:
            view_label = st.selectbox("Revenue view", list(VIEW_OPTIONS.keys()), key="pl_view_select")
            section = VIEW_OPTIONS[view_label]
            st.divider()
            all_clients = sorted(set(
                c for k in section if k in st.session_state["pipeline"]
                and isinstance(st.session_state["pipeline"][k], pd.DataFrame)
                for c in st.session_state["pipeline"][k]["Client"].tolist()
            ))
            client_options = ["All clients"] + all_clients
            selected_client = st.selectbox("Filter clients", client_options, index=0)
            selected_clients = all_clients if selected_client == "All clients" else [selected_client]
        else:
            section = None
            selected_clients = []
    else:
        section = None
        selected_clients = []


@st.dialog("Month Breakdown", width="large")
def show_month_breakdown(month, pipeline, section, selected_clients, token):
    st.subheader(f"{month} — Revenue by Client & Project")

    ALL_SECTIONS = [
        ("CONFIRMED",       "Confirmed Revenue",  "#E0EAF6"),
        ("PROPOSED",        "Proposed",           "#E4EFDC"),
        ("ACCOUNT PLANNING","Account Planning",   "#FDF3D0"),
        ("SPECULATIVE",     "Speculative",        "#F0908A"),
    ]

    # Get clients that have revenue this month for each section
    for sec_key, sec_label, sec_color in ALL_SECTIONS:
        if sec_key not in pipeline or not isinstance(pipeline[sec_key], pd.DataFrame):
            continue
        sec_df = pipeline[sec_key].copy()
        if selected_clients:
            sec_df = sec_df[sec_df["Client"].isin(selected_clients)]
        if month not in sec_df.columns:
            continue
        sec_df[month] = pd.to_numeric(sec_df[month], errors="coerce").fillna(0)
        active = sec_df[sec_df[month] > 0]
        if active.empty:
            continue

        st.markdown(
            f"<div style='background:{sec_color};padding:8px 12px;border-radius:8px;"
            f"font-weight:700;margin-top:16px'>{sec_label}</div>",
            unsafe_allow_html=True,
        )

        rows_out = []
        for _, client_row in active.iterrows():
            rows_out.append({
                "Client":  client_row["Client"],
                "Revenue": float(client_row[month]),
            })

        df_sec = pd.DataFrame(rows_out).sort_values("Revenue", ascending=False)
        total = df_sec["Revenue"].sum()
        st.caption(f"Total: **£{total:,.0f}**")
        df_sec["Revenue"] = df_sec["Revenue"].apply(lambda v: f"£{v:,.0f}")
        st.dataframe(df_sec, use_container_width=True, hide_index=True)

def _clean_val(v):
    try:
        return float(v) if v else 0.0
    except (TypeError, ValueError):
        return 0.0

# ── Main ──────────────────────────────────────────────────────────────────────
st.markdown(f"<h1 style='color:{NV_DARK}'>Neverland Dashboard</h1>", unsafe_allow_html=True)
tab_pipeline, tab_pl, tab_bva, tab_cap, tab_scope, tab_revenue, tab_sow = st.tabs(["Pipeline", "P&L", "Budget vs Actual", "Capacity", "Capacity vs Chargeout", "Revenue Tracker", "Capacity Planning"])

CATEGORY_COLORS = {
    "Confirmed Revenue": "#E0EAF6",
    "Proposed":          "#E4EFDC",
    "Account Planning":  "#FDF3D0",
    "Speculative":       "#F0908A",
}


# ── Pipeline Tab ──────────────────────────────────────────────────────────────
with tab_pipeline:
    if "pipeline" not in st.session_state or not section:
        st.info("Click **Sync from SharePoint** in the sidebar to load your pipeline.", icon="ℹ️")
    else:
        pipeline = st.session_state["pipeline"]
        frames = [pipeline[k].copy() for k in section if k in pipeline and isinstance(pipeline[k], pd.DataFrame)]
        if not frames:
            st.warning("No data found for selected view.")
        else:
            combined = pd.concat(frames, ignore_index=True)
            combined[MONTHS] = combined[MONTHS].apply(pd.to_numeric, errors="coerce").fillna(0)
            df = combined.groupby("Client", as_index=False)[MONTHS].sum()
            if selected_clients:
                df = df[df["Client"].isin(selected_clients)]

            total_annual = df[MONTHS].sum().sum()
            best_month   = df[MONTHS].sum().idxmax()
            best_month_v = df[MONTHS].sum().max()
            top_client   = df.set_index("Client")[MONTHS].sum(axis=1).idxmax()
            top_client_v = df.set_index("Client")[MONTHS].sum(axis=1).max()

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Total Annual Revenue", fmt_gbp(total_annual))
            k2.metric("Best Month",           f"{best_month}  {fmt_gbp(best_month_v)}")
            k3.metric("Top Client",           top_client)
            k4.metric("Top Client Revenue",   fmt_gbp(top_client_v))
            st.divider()

            chart_view = st.radio("Chart view", ["By Category", "By Client"], horizontal=True)

            cat_frames = []
            for k in section:
                if k not in pipeline or not isinstance(pipeline[k], pd.DataFrame):
                    continue
                sec_df = pipeline[k].copy()
                if selected_clients:
                    sec_df = sec_df[sec_df["Client"].isin(selected_clients)]
                sec_df[MONTHS] = sec_df[MONTHS].apply(pd.to_numeric, errors="coerce").fillna(0)
                cat_frames.append({"Category": SECTION_LABELS.get(k, k), **sec_df[MONTHS].sum().to_dict()})

            cat_df   = pd.DataFrame(cat_frames)
            cat_long = cat_df.melt(id_vars="Category", value_vars=MONTHS, var_name="Month", value_name="Revenue")
            cat_long["Month"] = pd.Categorical(cat_long["Month"], categories=MONTHS, ordered=True)
            cat_long = cat_long[cat_long["Revenue"] != 0]
            cat_long["Label"] = cat_long["Revenue"].apply(fmt_gbp)

            grand_totals = cat_long.groupby("Month", observed=True)["Revenue"].sum()
            month_labels = [f"{m}<br><b>£{grand_totals.get(m, 0):,.0f}</b>" for m in MONTHS]

            if chart_view == "By Category":
                fig_cat = px.bar(cat_long, x="Month", y="Revenue", color="Category", text="Label",
                    template="plotly_white", color_discrete_map=CATEGORY_COLORS,
                    labels={"Revenue": "Revenue (£)", "Month": ""}, height=520)
                fig_cat.update_traces(textposition="inside", textfont_size=10)
                fig_cat.update_layout(
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
                    xaxis=dict(tickvals=MONTHS, ticktext=month_labels),
                    yaxis_tickprefix="£", yaxis_tickformat=",.0f", bargap=0.2)
                sel = st.plotly_chart(fig_cat, use_container_width=True, on_select="rerun", selection_mode="points", key="cat_chart")
                if sel and sel.get("selection", {}).get("points"):
                    raw_x = str(sel["selection"]["points"][0].get("x", ""))
                    clicked_month = next((m for m in MONTHS if m in raw_x), None)
                    if clicked_month:
                        show_month_breakdown(clicked_month, pipeline, section, selected_clients, st.session_state.get("sp_token"))
            else:
                long = df.melt(id_vars="Client", value_vars=MONTHS, var_name="Month", value_name="Revenue")
                long["Month"] = pd.Categorical(long["Month"], categories=MONTHS, ordered=True)
                long = long[long["Revenue"] != 0]
                long["Label"] = long["Revenue"].apply(fmt_gbp)
                client_totals = long.groupby("Month", observed=True)["Revenue"].sum()
                client_month_labels = [f"{m}<br><b>£{client_totals.get(m, 0):,.0f}</b>" for m in MONTHS]
                fig_cli = px.bar(long, x="Month", y="Revenue", color="Client", text="Label",
                    template="plotly_white", color_discrete_sequence=px.colors.qualitative.Safe,
                    labels={"Revenue": "Revenue (£)", "Month": ""}, height=520)
                fig_cli.update_traces(textposition="inside", textfont_size=10)
                fig_cli.update_layout(
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
                    xaxis=dict(tickvals=MONTHS, ticktext=client_month_labels),
                    yaxis_tickprefix="£", yaxis_tickformat=",.0f", bargap=0.2)
                sel2 = st.plotly_chart(fig_cli, use_container_width=True, on_select="rerun", selection_mode="points", key="cli_chart")
                if sel2 and sel2.get("selection", {}).get("points"):
                    raw_x2 = str(sel2["selection"]["points"][0].get("x", ""))
                    clicked_month2 = next((m for m in MONTHS if m in raw_x2), None)
                    if clicked_month2:
                        show_month_breakdown(clicked_month2, pipeline, section, selected_clients, st.session_state.get("sp_token"))

            st.subheader("Monthly Totals")
            monthly = df[MONTHS].sum().reset_index()
            monthly.columns = ["Month", "Revenue"]
            monthly["Month"] = pd.Categorical(monthly["Month"], categories=MONTHS, ordered=True)
            monthly["Label"] = monthly["Revenue"].apply(fmt_gbp)
            fig2 = px.area(monthly, x="Month", y="Revenue", text="Label", template="plotly_white",
                labels={"Revenue": "Revenue (£)", "Month": ""}, height=320)
            fig2.update_traces(line_color=NV_PINK, fillcolor="rgba(102,126,234,0.15)",
                textposition="top center", textfont_size=11)
            fig2.update_layout(yaxis_tickprefix="£", yaxis_tickformat=",.0f")
            st.plotly_chart(fig2, use_container_width=True)

            st.subheader("Client Breakdown")
            summary = df.copy()
            summary["Total"] = summary[MONTHS].sum(axis=1)
            summary = summary.sort_values("Total", ascending=False)

            # Format client rows
            summary_display = summary.copy()
            for m in MONTHS + ["Total"]:
                summary_display[m] = summary_display[m].apply(fmt_gbp)

            def bold_total_col(row):
                return ["font-weight:bold" if c == "Total" else "" for c in row.index]

            st.dataframe(
                summary_display.style.apply(bold_total_col, axis=1),
                use_container_width=True, hide_index=True,
            )

            # Grand total pinned at bottom
            grand_vals = {m: summary[m].sum() for m in MONTHS}
            grand_vals["Total"] = summary["Total"].sum()
            grand_df = pd.DataFrame([{"Client": "Grand Total", **{m: fmt_gbp(v) for m, v in grand_vals.items()}}])
            st.markdown("""
            <style>
            div[data-testid="stDataFrame"]:last-of-type thead { display: none; }
            </style>""", unsafe_allow_html=True)
            st.dataframe(
                grand_df.style.set_properties(**{"font-weight": "bold", "background-color": "#f0f0f0"}),
                use_container_width=True, hide_index=True,
            )

            st.divider()
            st.subheader("All Categories — Annual Totals")
            overview_rows = []
            for sec, sec_df in pipeline.items():
                if not isinstance(sec_df, pd.DataFrame):
                    continue
                total = sec_df[MONTHS].sum().sum()
                overview_rows.append({"Category": SECTION_LABELS.get(sec, sec), "Total Revenue": total})
            overview = pd.DataFrame(overview_rows)
            fig3 = px.bar(overview, x="Category", y="Total Revenue", template="plotly_white",
                color="Category", color_discrete_map=CATEGORY_COLORS,
                text=overview["Total Revenue"].apply(fmt_gbp), height=320)
            fig3.update_traces(textposition="outside")
            fig3.update_layout(showlegend=False, yaxis_tickprefix="£", yaxis_tickformat=",.0f")
            st.plotly_chart(fig3, use_container_width=True)

# ── P&L Tab ───────────────────────────────────────────────────────────────────
with tab_pl:
    if "budget" not in st.session_state or "pipeline" not in st.session_state:
        st.info("Sync both Pipeline and Budget files to view the P&L.", icon="ℹ️")
    else:
        budget   = st.session_state["budget"]
        pipeline = st.session_state["pipeline"]

        pl_view = st.session_state.get("pl_view_select", "Confirmed")
        pl_keys = PL_VIEWS.get(pl_view, ["CONFIRMED"])

        # Income from pipeline
        pl_frames = [pipeline[k].copy() for k in pl_keys if k in pipeline and isinstance(pipeline[k], pd.DataFrame)]
        if pl_frames:
            pl_combined = pd.concat(pl_frames, ignore_index=True)
            pl_combined[MONTHS] = pl_combined[MONTHS].apply(pd.to_numeric, errors="coerce").fillna(0)
            income = [float(pl_combined[m].sum()) for m in MONTHS]
        else:
            income = [0.0] * 12

        if "__office_debug__" in budget:
            with st.expander("🔍 Jan Office Cost breakdown (debug)"):
                st.write("Detected month cols:", budget.get("__det_month_cols__"))
                st.write("Row 24 raw (first 16 cols):", budget.get("__row24_raw__"))
                st.write("Row 54 raw (first 16 cols):", budget.get("__row54_raw__"))
                st.dataframe(pd.DataFrame(budget["__office_debug__"]), use_container_width=True)
                st.metric("Jan Total after deduction", f"£{budget.get('__office_total_jan__', 0):,.0f}")

        # Costs from budget file
        staff      = [budget.get("Total staff costs", {}).get(m, 0) for m in MONTHS]
        office     = [budget.get("Total office & Other costs", {}).get(m, 0) for m in MONTHS]
        net_profit = [income[i] - staff[i] - office[i] for i in range(12)]
        margin     = [net_profit[i] / income[i] if income[i] else 0 for i in range(12)]

        # KPIs
        k1, k2, k3, k4 = st.columns(4)
        total_inc  = sum(income)
        total_np   = sum(net_profit)
        pct        = lambda v: f"{v/total_inc*100:.1f}%" if total_inc else "—"

        k1.metric("Total Income",       f"£{total_inc:,.0f}")
        k2.metric("Total Staff Costs",  f"£{sum(staff):,.0f}",      delta=pct(sum(staff)),   delta_color="off")
        k3.metric("Total Office Costs", f"£{sum(office):,.0f}",     delta=pct(sum(office)),  delta_color="off")
        k4.metric("Annual Net Profit",  f"£{total_np:,.0f}",        delta=pct(total_np),     delta_color="normal" if total_np >= 0 else "inverse")

        st.divider()

        st.subheader("P&L by Month")

        def style_pl_row(row):
            bold_rows = ["Total Income", "Net Profit", "Operating Profit Margin"]
            styles = []
            for val in row:
                s = str(val).replace("£", "").replace(",", "").replace("%", "").strip()
                base = "font-weight:bold" if row.name in bold_rows else ""
                try:
                    n = float(s)
                    if row.name in ["Net Profit", "Operating Profit Margin"]:
                        color = "color:#c0392b" if n < 0 else "color:#27ae60"
                        styles.append(f"{base}; {color}" if base else color)
                    else:
                        styles.append(base)
                except ValueError:
                    styles.append(base)
            return styles

        pl_table = pd.DataFrame({
            "":      ["Total Income", "Total Staff Costs", "Total Office & Other", "Net Profit", "Operating Profit Margin"],
            **{m: [
                f"£{income[i]:,.0f}",
                f"£{staff[i]:,.0f}",
                f"£{office[i]:,.0f}",
                f"£{net_profit[i]:,.0f}",
                f"{margin[i]*100:.0f}%",
            ] for i, m in enumerate(MONTHS)},
            "Total": [
                f"£{sum(income):,.0f}",
                f"£{sum(staff):,.0f}",
                f"£{sum(office):,.0f}",
                f"£{sum(net_profit):,.0f}",
                f"{(sum(net_profit)/sum(income)*100) if sum(income) else 0:.0f}%",
            ]
        }).set_index("")

        st.dataframe(
            pl_table.style.apply(style_pl_row, axis=1),
            use_container_width=True,
        )

# ── Budget vs Actual Tab ──────────────────────────────────────────────────────
with tab_bva:
    if "budget" not in st.session_state or "pipeline" not in st.session_state:
        st.info("Sync both Pipeline and Budget files to see Budget vs Actual.", icon="ℹ️")
    else:
        budget   = st.session_state["budget"]
        pipeline = st.session_state["pipeline"]

        bva_view = st.session_state.get("pl_view_select", "Confirmed")
        bva_keys = PL_VIEWS.get(bva_view, ["CONFIRMED"])

        actual_frames = [pipeline[k].copy() for k in bva_keys if k in pipeline and isinstance(pipeline[k], pd.DataFrame)]
        if actual_frames:
            actual_combined = pd.concat(actual_frames, ignore_index=True)
            actual_combined[MONTHS] = actual_combined[MONTHS].apply(pd.to_numeric, errors="coerce").fillna(0)
            actual_monthly = actual_combined[MONTHS].sum()
        else:
            actual_monthly = pd.Series({m: 0 for m in MONTHS})

        budget_monthly = pd.Series({m: budget.get("Total income", {}).get(m, 0) for m in MONTHS})

        variance       = actual_monthly - budget_monthly
        variance_pct   = ((actual_monthly / budget_monthly) * 100).fillna(0)

        # KPIs
        k1, k2, k3 = st.columns(3)
        k1.metric("Budget Total",  f"£{budget_monthly.sum():,.0f}")
        k2.metric("Actual Total",  f"£{actual_monthly.sum():,.0f}")
        var_total = actual_monthly.sum() - budget_monthly.sum()
        k3.metric("Variance",      f"£{var_total:,.0f}")

        st.divider()
        st.subheader("Budget vs Actual by Month")

        bva_df = pd.DataFrame({
            "Month":  MONTHS * 2,
            "Value":  list(budget_monthly) + list(actual_monthly),
            "Type":   (["Budget"] * 12) + ([bva_view] * 12),
        })
        bva_df["Month"] = pd.Categorical(bva_df["Month"], categories=MONTHS, ordered=True)
        bva_df["Label"] = bva_df["Value"].apply(lambda v: f"£{v:,.0f}")

        fig_bva = px.bar(bva_df, x="Month", y="Value", color="Type", barmode="group",
            text="Label", template="plotly_white", height=480,
            color_discrete_map={"Budget": "#E0EAF6", bva_view: NV_PINK},
            labels={"Value": "£", "Month": ""})
        fig_bva.update_traces(textposition="outside", textfont_size=9)
        fig_bva.update_layout(yaxis_tickprefix="£", yaxis_tickformat=",.0f",
            legend=dict(orientation="h", y=1.05))
        st.plotly_chart(fig_bva, use_container_width=True)

        st.subheader("Variance % vs Budget")
        var_df = pd.DataFrame({"Month": MONTHS, "Variance %": list(variance_pct - 100)})
        var_df["Month"] = pd.Categorical(var_df["Month"], categories=MONTHS, ordered=True)
        fig_var = px.bar(var_df, x="Month", y="Variance %", template="plotly_white", height=300,
            color="Variance %", color_continuous_scale=["#F0908A", "#ffffff", "#E4EFDC"],
            labels={"Variance %": "% vs Budget", "Month": ""})
        fig_var.update_layout(yaxis_ticksuffix="%", coloraxis_showscale=False)
        fig_var.add_hline(y=0, line_dash="dash", line_color="grey")
        st.plotly_chart(fig_var, use_container_width=True)

        st.subheader("Monthly Detail")
        detail = pd.DataFrame({
            "Month":       MONTHS,
            "Budget":      [f"£{v:,.0f}" for v in budget_monthly],
            "Actual":      [f"£{v:,.0f}" for v in actual_monthly],
            "Variance":    [f"£{v:,.0f}" for v in variance],
            "vs Budget %": [f"{v:.0f}%" for v in (variance_pct - 100)],
        })
        st.dataframe(detail, use_container_width=True, hide_index=True)

# ── Capacity Tab ──────────────────────────────────────────────────────────────
with tab_cap:
    if "capacity" not in st.session_state:
        st.info("Click **Sync Budget File** in the sidebar to load capacity data.", icon="ℹ️")
    else:
        cap_df = st.session_state["capacity"]
        cap_df["Month"] = pd.Categorical(cap_df["Month"], categories=MONTHS, ordered=True)
        cap_df["Cost"]  = pd.to_numeric(cap_df["Cost"], errors="coerce").fillna(0)

        departments = sorted(cap_df["Department"].unique())

        # KPIs
        total_cost   = cap_df["Cost"].sum()
        total_people = cap_df.groupby(["Department", "Employee"]).first().reset_index()["Employee"].nunique()
        k1, k2, k3   = st.columns(3)
        k1.metric("Total Annual Cost",  f"£{total_cost:,.0f}")
        k2.metric("Departments",        len(departments))
        k3.metric("People",             total_people)

        st.divider()

        # Stacked bar by department per month
        st.subheader("Cost by Department & Month")
        dept_monthly = cap_df.groupby(["Department", "Month"], observed=True)["Cost"].sum().reset_index()
        dept_monthly["Label"] = dept_monthly["Cost"].apply(fmt_gbp)

        fig_cap = px.bar(
            dept_monthly, x="Month", y="Cost", color="Department",
            text="Label", template="plotly_white", height=500,
            color_discrete_sequence=px.colors.qualitative.Safe,
            labels={"Cost": "Cost (£)", "Month": ""},
        )
        fig_cap.update_traces(textposition="inside", textfont_size=10)
        month_totals = dept_monthly.groupby("Month", observed=True)["Cost"].sum()
        cap_month_labels = [f"{m}<br><b>£{month_totals.get(m, 0):,.0f}</b>" for m in MONTHS]
        fig_cap.update_layout(
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
            xaxis=dict(tickvals=MONTHS, ticktext=cap_month_labels),
            yaxis_tickprefix="£", yaxis_tickformat=",.0f", bargap=0.2,
        )
        st.plotly_chart(fig_cap, use_container_width=True)

        # Department summary table with job titles + bold subtotals
        st.subheader("Department Summary")
        dept_staff_filter = st.selectbox("Staff type", ["All", "Full Time", "Freelancer"], key="dept_staff_filter")
        dept_cap_df = cap_df.copy()
        if dept_staff_filter != "All" and "Staff Type" in dept_cap_df.columns:
            dept_cap_df = dept_cap_df[dept_cap_df["Staff Type"] == dept_staff_filter]

        dept_monthly_totals = dept_cap_df.groupby(["Department", "Month"], observed=True)["Cost"].sum().reset_index()
        dept_pivot = dept_monthly_totals.pivot(index="Department", columns="Month", values="Cost").fillna(0)
        dept_pivot = dept_pivot.reindex(columns=MONTHS, fill_value=0)
        dept_pivot["Total"] = dept_pivot[MONTHS].sum(axis=1)
        dept_pivot = dept_pivot.sort_values("Total", ascending=False)
        dept_pivot.index.name = "Department"

        # Add grand total row
        grand_total = pd.DataFrame(
            [{m: dept_pivot[m].sum() for m in MONTHS + ["Total"]}],
            index=["Grand Total"]
        )
        grand_total.index.name = "Department"
        dept_with_total = pd.concat([dept_pivot, grand_total])

        for col in MONTHS + ["Total"]:
            dept_with_total[col] = dept_with_total[col].apply(lambda v: f"£{v:,.0f}")

        def style_dept(row):
            if row.name == "Grand Total":
                return ["font-weight:bold; background-color:#f0f0f0"] * len(row)
            return ["font-weight:bold" if c == "Total" else "" for c in row.index]

        st.dataframe(
            dept_with_total.style.apply(style_dept, axis=1),
            use_container_width=True,
        )

        # Cost breakdown — by Job Title or by Person
        st.divider()
        fc1, fc2 = st.columns(2)
        with fc1:
            selected_dept = st.selectbox("Select department", ["All"] + departments, key="cap_dept")
        with fc2:
            breakdown_type = st.selectbox("View by", ["Job Title", "Person"], key="cap_breakdown")

        filt_df = cap_df.copy()
        if selected_dept != "All":
            filt_df = filt_df[filt_df["Department"] == selected_dept]
        group_col = "Job Title" if breakdown_type == "Job Title" else "Employee"
        st.subheader(f"Cost by {breakdown_type}")

        # Build table grouped by department with subtotals and gaps
        table_rows   = []
        subtotal_idx = []
        empty_idx    = []
        grand_idx    = []
        row_idx      = 0

        dept_list = sorted(filt_df["Department"].unique())
        for di, dept in enumerate(dept_list):
            dept_data = filt_df[filt_df["Department"] == dept]
            grp = dept_data.groupby([group_col, "Month"], observed=True)["Cost"].sum().reset_index()
            grp = grp.pivot(index=group_col, columns="Month", values="Cost").fillna(0)
            grp = grp.reindex(columns=MONTHS, fill_value=0)
            grp["Total"] = grp[MONTHS].sum(axis=1)
            grp = grp.sort_values("Total", ascending=False)

            for name, row in grp.iterrows():
                table_rows.append({"": name, **{m: row[m] for m in MONTHS}, "Total": row["Total"]})
                row_idx += 1

            # Subtotal row
            sub = {m: dept_data[dept_data["Month"] == m]["Cost"].sum() for m in MONTHS}
            table_rows.append({"": f"{dept} — Total", **sub, "Total": sum(sub.values())})
            subtotal_idx.append(row_idx)
            row_idx += 1

            # Gap row (empty) between departments
            if di < len(dept_list) - 1:
                table_rows.append({"": "", **{m: "" for m in MONTHS}, "Total": ""})
                empty_idx.append(row_idx)
                row_idx += 1

        # Grand total
        all_costs = {m: filt_df[filt_df["Month"] == m]["Cost"].sum() for m in MONTHS}
        table_rows.append({"": "Grand Total", **all_costs, "Total": sum(all_costs.values())})
        grand_idx.append(row_idx)

        result_df = pd.DataFrame(table_rows)
        # Use numeric index to avoid duplicate key errors with Styler
        result_df.index = range(len(result_df))

        # Format numbers (skip empty gap rows)
        for col in MONTHS + ["Total"]:
            result_df[col] = result_df[col].apply(
                lambda v: f"£{v:,.0f}" if isinstance(v, (int, float)) else v
            )

        def style_rows(row):
            if row.name in grand_idx:
                return ["font-weight:bold; background-color:#e8e8e8"] * len(row)
            if row.name in subtotal_idx:
                return ["font-weight:bold; background-color:#f5f5f5"] * len(row)
            if row.name in empty_idx:
                return ["background-color:#ffffff"] * len(row)
            return ["font-weight:bold" if c == "Total" else "" for c in row.index]

        st.dataframe(
            result_df.style.apply(style_rows, axis=1),
            use_container_width=True,
            hide_index=True,
        )

# ── Capacity vs Chargeout Tab ─────────────────────────────────────────────────
with tab_scope:
    if "capacity" not in st.session_state or "scope" not in st.session_state:
        st.info("Sync Budget File to load Capacity vs Chargeout data.", icon="ℹ️")
        token = st.session_state.get("sp_token")
        if token and st.button("🔍 Debug scope structure", key="scope_struct_btn"):
            try:
                rows = fetch_scope_structure_debug(token)
                st.dataframe(pd.DataFrame(rows), use_container_width=True)
            except Exception as e:
                st.error(str(e))
    else:
        cap_df   = st.session_state["capacity"]
        scope_df = st.session_state["scope"]

        if scope_df.empty:
            st.warning("No scope data found.")
        else:
            scope_df["Month"]     = pd.Categorical(scope_df["Month"], categories=MONTHS, ordered=True)
            scope_df["Chargeout"] = pd.to_numeric(scope_df["Chargeout"], errors="coerce").fillna(0)
            cap_df["Month"]       = pd.Categorical(cap_df["Month"], categories=MONTHS, ordered=True)
            cap_df["Cost"]        = pd.to_numeric(cap_df["Cost"], errors="coerce").fillna(0)

            QUARTERS = {
                "All Year": MONTHS,
                "Q1 (Jan-Mar)": ["Jan", "Feb", "Mar"],
                "Q2 (Apr-Jun)": ["Apr", "May", "Jun"],
                "Q3 (Jul-Sep)": ["Jul", "Aug", "Sep"],
                "Q4 (Oct-Dec)": ["Oct", "Nov", "Dec"],
            }

            sc1, sc2, sc3 = st.columns(3)
            with sc1:
                quarter = st.selectbox("Quarter", list(QUARTERS.keys()), key="scope_quarter")
            with sc2:
                scope_staff = st.selectbox("Staff type", ["All", "Full Time", "Freelancer"], key="scope_staff")
            with sc3:
                scope_cat = st.selectbox("Category", ["All", "Client", "New Biz", "Client (not recovered)"], key="scope_cat")

            q_months = QUARTERS[quarter]

            # Filter capacity
            cap_filt = cap_df[cap_df["Month"].isin(q_months)]
            if scope_staff != "All" and "Staff Type" in cap_filt.columns:
                cap_filt = cap_filt[cap_filt["Staff Type"] == scope_staff]
            cap_by_dept = cap_filt.groupby("Department")["Cost"].sum().reset_index()
            cap_by_dept.columns = ["Department", "Capacity Cost"]

            # Filter scope
            scope_filt = scope_df[scope_df["Month"].isin(q_months)]
            if scope_cat != "All" and "Category" in scope_filt.columns:
                scope_filt = scope_filt[scope_filt["Category"] == scope_cat]
            scope_by_dept = scope_filt.groupby("Department")["Chargeout"].sum().reset_index()
            scope_by_dept.columns = ["Department", "Chargeout"]

            # Merge
            merged = pd.merge(cap_by_dept, scope_by_dept, on="Department", how="outer").fillna(0)
            merged["Variance"]    = merged["Chargeout"] - merged["Capacity Cost"]
            merged["Utilisation"] = (merged["Chargeout"] / merged["Capacity Cost"] * 100).replace([float("inf"), float("nan")], 0)

            # KPIs
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Total Capacity Cost",  f"£{merged['Capacity Cost'].sum():,.0f}")
            k2.metric("Total Chargeout",       f"£{merged['Chargeout'].sum():,.0f}")
            k3.metric("Variance",              f"£{merged['Variance'].sum():,.0f}")
            k4.metric("Overall Utilisation",   f"{merged['Chargeout'].sum() / merged['Capacity Cost'].sum() * 100:.0f}%" if merged['Capacity Cost'].sum() else "0%")

            st.divider()

            st.subheader(f"Capacity vs Chargeout — {quarter}")

            CAT_COLORS = {
                "Client":                NV_PINK,
                "New Biz":               "#F0908A",
                "Client (not recovered)":"#FDF3D0",
                "Capacity Cost":         "#E0EAF6",
            }

            # Pipeline monthly totals for the selected view
            if "pipeline" in st.session_state:
                pipe_frames = [st.session_state["pipeline"][k].copy()
                               for k in VIEW_OPTIONS.get(st.session_state.get("pl_view_select","Confirmed"), ["CONFIRMED"])
                               if k in st.session_state["pipeline"] and isinstance(st.session_state["pipeline"][k], pd.DataFrame)]
                if pipe_frames:
                    pipe_combined = pd.concat(pipe_frames, ignore_index=True)
                    pipe_combined[MONTHS] = pipe_combined[MONTHS].apply(pd.to_numeric, errors="coerce").fillna(0)
                    pipe_monthly_totals = pipe_combined[MONTHS].sum()
                else:
                    pipe_monthly_totals = pd.Series({m: 0 for m in MONTHS})
            else:
                pipe_monthly_totals = pd.Series({m: 0 for m in MONTHS})

            def make_stacked_vs_capacity(x_col, cap_data, scope_data, x_vals, pipe_data=None):
                fig = go.Figure()
                # Capacity as a line
                fig.add_trace(go.Scatter(
                    x=cap_data[x_col], y=cap_data["£"],
                    name="Capacity Cost", mode="lines+markers",
                    line=dict(color="#8899BB", width=2, dash="dot"),
                    marker=dict(size=6, color="#8899BB"),
                ))
                # Pipeline bar
                if pipe_data is not None:
                    fig.add_trace(go.Bar(
                        x=list(pipe_data.index), y=list(pipe_data.values),
                        name="Pipeline", marker_color=NV_PINK,
                        offsetgroup="C",
                        opacity=0.7,
                    ))
                # Stacked chargeout bars
                for cat in ["Client", "New Biz", "Client (not recovered)"]:
                    d = scope_data[scope_data["Type"] == cat]
                    if d.empty:
                        continue
                    fig.add_trace(go.Bar(
                        x=d[x_col], y=d["£"],
                        name=cat, marker_color=CAT_COLORS.get(cat, "#aaa"),
                        offsetgroup="B",
                    ))
                fig.update_layout(
                    barmode="stack", template="plotly_white", height=480,
                    yaxis_tickprefix="£", yaxis_tickformat=",.0f",
                    legend=dict(orientation="h", y=1.05),
                    bargroupgap=0.15,
                )
                return fig

            # Build chargeout by dept + category
            scope_by_cat = scope_filt.groupby(["Department", "Category"])["Chargeout"].sum().reset_index()
            scope_by_cat.columns = ["Department", "Type", "£"]
            cap_chart = cap_by_dept.rename(columns={"Capacity Cost": "£"}).assign(Type="Capacity Cost")
            if "Capacity Cost" in cap_chart.columns:
                cap_chart = cap_by_dept.copy()
                cap_chart.columns = ["Department", "£"]
                cap_chart["Type"] = "Capacity Cost"

            fig_scope = make_stacked_vs_capacity("Department", cap_chart, scope_by_cat, merged["Department"].tolist())
            st.plotly_chart(fig_scope, use_container_width=True)

            # % of capacity per department
            dept_pct = merged.copy()
            dept_pct["Chargeout % of Capacity"] = dept_pct.apply(
                lambda r: f"{r['Chargeout']/r['Capacity Cost']*100:.0f}%" if r["Capacity Cost"] else "—", axis=1
            )
            dept_pct_display = dept_pct[["Department", "Chargeout % of Capacity"]].set_index("Department").T

            def style_dept_pct(row):
                styles = []
                for val in row:
                    try:
                        n = float(str(val).replace("%",""))
                        if n >= 100:
                            styles.append("color:#27ae60; font-weight:bold")
                        elif n >= 75:
                            styles.append("color:#f39c12; font-weight:bold")
                        else:
                            styles.append("color:#c0392b; font-weight:bold")
                    except (ValueError, TypeError):
                        styles.append("")
                return styles

            st.dataframe(dept_pct_display.style.apply(style_dept_pct, axis=1), use_container_width=True)

            # Monthly breakdown chart
            st.subheader("Monthly Breakdown by Department")
            cap_monthly   = cap_filt.groupby(["Department", "Month"], observed=True)["Cost"].sum().reset_index()
            cap_monthly["Type"] = "Capacity Cost"
            cap_monthly = cap_monthly.rename(columns={"Cost": "£"})

            scope_monthly = scope_filt.groupby(["Department", "Month", "Category"], observed=True)["Chargeout"].sum().reset_index()
            scope_monthly = scope_monthly.rename(columns={"Chargeout": "£", "Category": "Type"})

            monthly_combined = pd.concat([cap_monthly, scope_monthly])

            dept_filter = st.selectbox("Filter department", ["All"] + sorted(merged["Department"].tolist()), key="scope_dept")
            if dept_filter != "All":
                monthly_combined = monthly_combined[monthly_combined["Department"] == dept_filter]

            plot_df = monthly_combined.copy()
            if dept_filter != "All":
                plot_df = plot_df[plot_df["Department"] == dept_filter]
            plot_df = plot_df.groupby(["Month", "Type"], observed=True)["£"].sum().reset_index()
            plot_df["Month"] = pd.Categorical(plot_df["Month"], categories=q_months, ordered=True)

            cap_monthly_plot   = plot_df[plot_df["Type"] == "Capacity Cost"]
            scope_monthly_plot = plot_df[plot_df["Type"] != "Capacity Cost"]
            pipe_plot = pipe_monthly_totals[pipe_monthly_totals.index.isin(q_months)]
            fig_m = make_stacked_vs_capacity("Month", cap_monthly_plot, scope_monthly_plot, q_months, pipe_data=pipe_plot)
            fig_m.update_layout(height=450)
            st.plotly_chart(fig_m, use_container_width=True)

            # ── % table ──────────────────────────────────────────────────────
            cap_by_month   = cap_monthly_plot.set_index("Month")["£"]
            scope_by_month = scope_monthly_plot.groupby("Month", observed=True)["£"].sum()
            pipe_by_month  = pipe_plot

            pct_rows = {"Metric": ["Pipeline % of Capacity", "Chargeout % of Capacity"]}
            for m in q_months:
                cap_v   = cap_by_month.get(m, 0)
                scope_v = scope_by_month.get(m, 0)
                pipe_v  = pipe_by_month.get(m, 0)
                pct_rows[m] = [
                    f"{pipe_v/cap_v*100:.0f}%" if cap_v else "—",
                    f"{scope_v/cap_v*100:.0f}%" if cap_v else "—",
                ]

            pct_df = pd.DataFrame(pct_rows).set_index("Metric")

            def style_pct(row):
                styles = []
                for val in row:
                    try:
                        n = float(str(val).replace("%",""))
                        if n >= 100:
                            styles.append("color:#27ae60; font-weight:bold")
                        elif n >= 75:
                            styles.append("color:#f39c12; font-weight:bold")
                        else:
                            styles.append("color:#c0392b; font-weight:bold")
                    except (ValueError, TypeError):
                        styles.append("")
                return styles

            st.dataframe(pct_df.style.apply(style_pct, axis=1), use_container_width=True)

            # Summary table
            st.subheader("Summary Table")

            t1, t2, t3 = st.columns(3)
            with t1:
                from_month = st.selectbox("From month", MONTHS, index=0, key="table_from")
            with t2:
                to_month = st.selectbox("To month", MONTHS, index=11, key="table_to")
            with t3:
                table_cat = st.selectbox("Chargeout category", ["All", "Client", "New Biz", "Client (not recovered)"], key="table_cat")

            from_idx  = MONTHS.index(from_month)
            to_idx    = MONTHS.index(to_month)
            if to_idx < from_idx:
                to_idx = from_idx
            table_months = MONTHS[from_idx:to_idx + 1]

            # Recalculate for selected month range
            cap_range  = cap_filt[cap_filt["Month"].isin(table_months)].groupby("Department")["Cost"].sum().reset_index()
            cap_range.columns = ["Department", "Capacity Cost"]

            scope_table = scope_filt[scope_filt["Month"].isin(table_months)]
            if table_cat != "All" and "Category" in scope_table.columns:
                scope_table = scope_table[scope_table["Category"] == table_cat]
            scope_range = scope_table.groupby("Department")["Chargeout"].sum().reset_index()
            merged_range = pd.merge(cap_range, scope_range, on="Department", how="outer").fillna(0)

            # Add salary costs by department
            if "salary_dept" in st.session_state and not st.session_state["salary_dept"].empty:
                sal_df = st.session_state["salary_dept"].copy()
                sal_df["Month"] = sal_df["Month"].astype(str)
                sal_range = sal_df[sal_df["Month"].isin(table_months)].groupby("Department")["Salary"].sum().reset_index()
                sal_range.columns = ["Department", "Total Staff Costs"]
                merged_range = pd.merge(merged_range, sal_range, on="Department", how="left").fillna(0)
            else:
                merged_range["Total Staff Costs"] = 0
            merged_range["Variance"]          = merged_range["Chargeout"] - merged_range["Total Staff Costs"]
            merged_range["Target Multiplier"] = (merged_range["Capacity Cost"] / merged_range["Total Staff Costs"]).replace([float("inf"), float("nan")], 0)
            merged_range["Multiplier"]        = (merged_range["Chargeout"] / merged_range["Total Staff Costs"]).replace([float("inf"), float("nan")], 0)
            merged_range["Utilisation"] = (merged_range["Chargeout"] / merged_range["Capacity Cost"] * 100).replace([float("inf"), float("nan")], 0)

            display = merged_range.copy()

            # Add bold grand total row
            total_cap  = merged_range["Capacity Cost"].sum()
            total_sal  = merged_range["Total Staff Costs"].sum()
            total_char = merged_range["Chargeout"].sum()
            total_var  = total_char - total_sal
            total_util = (total_char / total_cap * 100) if total_cap else 0
            total_row  = pd.DataFrame([{
                "Department":        "Grand Total",
                "Capacity Cost":     total_cap,
                "Total Staff Costs": total_sal,
                "Chargeout":         total_char,
                "Variance":          total_var,
                "Multiplier":         total_char / total_sal if total_sal else 0,
                "Utilisation":       total_util,
            }])
            display = pd.concat([display, total_row], ignore_index=True)

            # Rename columns for clarity
            display = display.rename(columns={
                "Variance": "Gross Profit",
            })
            # Reorder columns
            col_order = ["Department", "Capacity Cost", "Total Staff Costs", "Chargeout", "Gross Profit", "Utilisation"]
            display = display[[c for c in col_order if c in display.columns]]

            # Store raw numeric values for colour coding before formatting
            gross_profit_raw = display["Gross Profit"].copy()

            # Merge Target Multiplier into Capacity Cost column
            raw_cap = merged_range.set_index("Department")["Capacity Cost"].to_dict()
            raw_sal = merged_range.set_index("Department")["Total Staff Costs"].to_dict()
            raw_cap["Grand Total"] = total_cap
            raw_sal["Grand Total"] = total_sal

            def fmt_cap(row):
                dept = row["Department"]
                cap = raw_cap.get(dept, 0)
                sal = raw_sal.get(dept, 0)
                mult = f" ({cap/sal:.2f}x)" if sal else ""
                return f"£{cap:,.0f}{mult}"
            display["Capacity Cost"] = display.apply(fmt_cap, axis=1)
            display["Total Staff Costs"] = display["Total Staff Costs"].apply(lambda v: f"£{v:,.0f}")
            display["Chargeout"]         = display["Chargeout"].apply(lambda v: f"£{v:,.0f}")
            raw_mult = merged_range.set_index("Department")["Multiplier"].to_dict()
            raw_mult["Grand Total"] = total_char / total_sal if total_sal else 0

            def fmt_gp(row):
                dept = row["Department"]
                gp   = gross_profit_raw.iloc[display.index.get_loc(row.name)] if row.name in display.index else row["Gross Profit"]
                mult = raw_mult.get(dept, 0)
                mult_str = f" ({mult:.2f}x)" if mult else ""
                try:
                    return f"£{float(gp):,.0f}{mult_str}"
                except (ValueError, TypeError):
                    return f"{gp}{mult_str}"
            display["Gross Profit"] = display.apply(fmt_gp, axis=1)
            display["Utilisation"]       = display["Utilisation"].apply(lambda v: f"{v:.0f}%")

            def bold_total(row):
                if row["Department"] == "Grand Total":
                    return ["font-weight:bold; background-color:#f0f0f0"] * len(row)
                # Colour Gross Profit red/green
                styles = []
                for c in row.index:
                    if c == "Gross Profit":
                        try:
                            raw = gross_profit_raw.iloc[display.index.get_loc(row.name)] if row.name in display.index else 0
                            styles.append("color:#27ae60; font-weight:bold" if raw >= 0 else "color:#c0392b; font-weight:bold")
                        except Exception:
                            styles.append("")
                    else:
                        styles.append("")
                return styles

            st.caption("💡 Click a row to see the breakdown by person")
            sel_summary = st.dataframe(
                display.style.apply(bold_total, axis=1),
                use_container_width=True, hide_index=True,
                on_select="rerun", selection_mode="single-row", key="summary_sel"
            )

            # Drilldown by person for selected department
            if sel_summary and sel_summary.get("selection", {}).get("rows"):
                row_idx  = sel_summary["selection"]["rows"][0]
                sel_dept = display.iloc[row_idx]["Department"]
                if sel_dept != "Grand Total" and "capacity" in st.session_state:
                    cap_df2   = st.session_state["capacity"]
                    scope_df2 = st.session_state.get("scope", pd.DataFrame())

                    st.markdown(f"### 👤 {sel_dept} — Breakdown by Person ({table_months[0]}–{table_months[-1]})")

                    # Capacity by person
                    dept_cap = cap_df2[
                        (cap_df2["Department"] == sel_dept) &
                        (cap_df2["Month"].isin(table_months))
                    ]
                    if "Staff Type" in dept_cap.columns and scope_staff != "All":
                        dept_cap = dept_cap[dept_cap["Staff Type"] == scope_staff]
                    cap_person = dept_cap.groupby("Employee")["Cost"].sum().reset_index()
                    cap_person.columns = ["Person", "Capacity Cost"]
                    cap_person["Capacity Cost"] = cap_person["Capacity Cost"].apply(lambda v: f"£{v:,.0f}")

                    # Chargeout by person
                    if not scope_df2.empty and "Category" in scope_df2.columns:
                        dept_scope = scope_df2[
                            (scope_df2["Department"] == sel_dept) &
                            (scope_df2["Month"].isin(table_months))
                        ]
                        if scope_cat != "All":
                            dept_scope = dept_scope[dept_scope["Category"] == scope_cat]
                        scope_person = dept_scope.groupby("Employee")["Chargeout"].sum().reset_index()
                        scope_person.columns = ["Person", "Chargeout"]
                        scope_person["Chargeout"] = scope_person["Chargeout"].apply(lambda v: f"£{v:,.0f}")
                        person_df = pd.merge(cap_person, scope_person, on="Person", how="outer").fillna("£0")
                    else:
                        person_df = cap_person

                    st.dataframe(person_df, use_container_width=True, hide_index=True)

# ── Revenue Tracker Tab ───────────────────────────────────────────────────────
with tab_revenue:
    if "pipeline" not in st.session_state or not section:
        st.info("Click **Sync from SharePoint** in the sidebar to load your pipeline.", icon="ℹ️")
    else:
        import datetime as _dt
        pipeline = st.session_state["pipeline"]

        cur_month_name = _dt.date.today().strftime("%b")
        if cur_month_name not in MONTHS:
            cur_month_name = MONTHS[0]
        cur_month_idx = MONTHS.index(cur_month_name)
        cur_quarter_months = MONTHS[(cur_month_idx // 3) * 3 : (cur_month_idx // 3) * 3 + 3]

        PERIOD_OPTIONS = {
            "Current Month":   [cur_month_name],
            "Current Quarter": cur_quarter_months,
            "Full Year":       MONTHS,
        }

        rv_keys = section  # driven by sidebar "Revenue view" selector

        rv_period = st.selectbox(
            "Period",
            list(PERIOD_OPTIONS.keys()),
            index=2,
            key="rv_period_select",
        )
        rv_months = PERIOD_OPTIONS[rv_period]

        rv_frames = [
            pipeline[k].copy()
            for k in rv_keys
            if k in pipeline and isinstance(pipeline[k], pd.DataFrame)
        ]
        if not rv_frames:
            st.warning("No data found for selected revenue type.")
            st.stop()

        rv_combined = pd.concat(rv_frames, ignore_index=True)
        rv_combined[MONTHS] = rv_combined[MONTHS].apply(pd.to_numeric, errors="coerce").fillna(0)
        rv_by_client = rv_combined.groupby("Client", as_index=False)[MONTHS].sum()
        if selected_clients:
            rv_by_client = rv_by_client[rv_by_client["Client"].isin(selected_clients)]
        rv_by_client["Total"] = rv_by_client[rv_months].sum(axis=1)
        rv_by_client = rv_by_client[rv_by_client["Total"] > 0].sort_values("Total", ascending=False).reset_index(drop=True)

        # ── Client summary (Total only) ───────────────────────────────────────
        st.subheader("Revenue by Client")
        display_rv = rv_by_client[["Client", "Total"]].copy()
        display_rv["Total"] = display_rv["Total"].apply(fmt_gbp)

        sel_rv = st.dataframe(
            display_rv,
            use_container_width=True,
            hide_index=True,
            on_select="rerun",
            selection_mode="single-row",
            key="rv_client_sel",
        )


        # ── Comparison ────────────────────────────────────────────────────────
        st.divider()

        QUARTER_MONTHS = {
            "Q1 (Jan–Mar)": ["Jan", "Feb", "Mar"],
            "Q2 (Apr–Jun)": ["Apr", "May", "Jun"],
            "Q3 (Jul–Sep)": ["Jul", "Aug", "Sep"],
            "Q4 (Oct–Dec)": ["Oct", "Nov", "Dec"],
        }

        cmp_mode = st.radio(
            "Compare by",
            ["Week", "Month", "Quarter", "Full Year"],
            horizontal=True,
            key="cmp_mode",
        )

        def _render_comparison(totals_a, totals_b, col_a, col_b, table_key="cmp_table"):
            all_clients = sorted(set(totals_a.index) | set(totals_b.index))
            cmp_rows = []
            for client in all_clients:
                val_a = totals_a.get(client, 0)
                val_b = totals_b.get(client, 0)
                change = val_a - val_b
                cmp_rows.append({
                    "Client": client, col_a: val_a, col_b: val_b,
                    "Change (£)": change,
                    "_change_raw": change,
                    "_new":  val_b == 0 and val_a > 0,
                    "_lost": val_a == 0 and val_b > 0,
                })
            cmp_df = pd.DataFrame(cmp_rows).sort_values("Change (£)", ascending=False)
            total_change = cmp_df["Change (£)"].sum()
            new_clients  = cmp_df[cmp_df["_new"]]["Client"].tolist()
            lost_clients = cmp_df[cmp_df["_lost"]]["Client"].tolist()
            def _fmt_delta(v):
                sign = "+" if v >= 0 else "-"
                av = abs(v)
                if av >= 1_000_000:
                    return f"{sign}£{av/1_000_000:.2f}M"
                if av >= 1_000:
                    return f"{sign}£{av/1_000:.0f}k"
                return f"{sign}£{av:,.0f}"
            ck1, ck2, ck3 = st.columns(3)
            ck1.metric(col_a, fmt_gbp(cmp_df[col_a].sum()), delta=_fmt_delta(total_change))
            ck2.metric(col_b, fmt_gbp(cmp_df[col_b].sum()))
            ck3.metric("Clients changed", len(cmp_df[cmp_df["Change (£)"] != 0]))
            if new_clients:
                st.success(f"New: {', '.join(new_clients)}")
            if lost_clients:
                st.warning(f"No longer active: {', '.join(lost_clients)}")
            display_cmp = cmp_df[["Client", col_a, col_b, "Change (£)"]].copy()
            display_cmp[col_a]        = display_cmp[col_a].apply(lambda v: fmt_gbp(v) if v else "—")
            display_cmp[col_b]        = display_cmp[col_b].apply(lambda v: fmt_gbp(v) if v else "—")
            display_cmp["Change (£)"] = display_cmp["Change (£)"].apply(lambda v: _fmt_delta(v) if v != 0 else "—")
            def _style_cmp(row):
                raw = cmp_df.loc[cmp_df["Client"] == row["Client"], "_change_raw"]
                cv  = raw.iloc[0] if not raw.empty else 0
                return ["color:#27ae60;font-weight:bold" if cv > 0 else ("color:#c0392b;font-weight:bold" if cv < 0 else "") if c == "Change (£)" else "" for c in row.index]
            sel = st.dataframe(
                display_cmp.style.apply(_style_cmp, axis=1),
                use_container_width=True, hide_index=True,
                on_select="rerun", selection_mode="single-row",
                key=table_key,
            )
            return cmp_df, sel

        # ── Period selector for Month / Quarter modes ─────────────────────────
        if cmp_mode == "Month":
            period_months = [st.selectbox("Month", MONTHS, index=MONTHS.index(cur_month_name) if cur_month_name in MONTHS else 0, key="cmp_month_sel")]
            subheader_label = f"Comparing {period_months[0]} — {rv_period}"
        elif cmp_mode == "Quarter":
            q_sel = st.selectbox("Quarter", list(QUARTER_MONTHS.keys()), index=0, key="cmp_q_sel")
            period_months = QUARTER_MONTHS[q_sel]
            subheader_label = f"Comparing {q_sel} — {rv_period}"
        elif cmp_mode == "Full Year":
            period_months = MONTHS
            subheader_label = f"Full Year Comparison — {rv_period}"
        else:
            period_months = rv_months
            subheader_label = f"Week-on-Week Comparison — {rv_period}"

        st.subheader(subheader_label)

        # ── Snapshot picker (shared by all modes) ─────────────────────────────
        if "snapshot_list" not in st.session_state:
            try:
                st.session_state["snapshot_list"] = list_pipeline_snapshots()
            except Exception:
                st.session_state["snapshot_list"] = []
        snapshots = st.session_state["snapshot_list"]

        def _fmt_snap(date_str):
            try:
                import datetime as _dt2
                d = _dt2.date.fromisoformat(date_str)
                day = d.day
                suffix = "th" if 11 <= (day % 100) <= 13 else {1:"st",2:"nd",3:"rd"}.get(day % 10, "th")
                return f"{day}{suffix} {d.strftime('%b %Y')}"
            except Exception:
                return date_str

        snap_label_a = "Compare date" if cmp_mode != "Week" else "Compare week"
        snap_label_b = "With date"    if cmp_mode != "Week" else "With week"

        if len(snapshots) < 1:
            st.info("No snapshots yet. Snapshots are saved automatically every Thursday at 6pm — check back after the first run.", icon="📸")
        else:
            snap_labels = {s: _fmt_snap(s) for s in snapshots}
            cmp1, cmp2 = st.columns(2)
            with cmp1:
                snap_a_date = st.selectbox(snap_label_a, snapshots, index=0, key="snap_a", format_func=lambda s: snap_labels[s])
            with cmp2:
                snap_b_options = [s for s in snapshots if s != snap_a_date]
                if snap_b_options:
                    snap_b_date = st.selectbox(snap_label_b, snap_b_options, index=0, key="snap_b", format_func=lambda s: snap_labels[s])
                else:
                    snap_b_date = None
                    st.info("Save at least 2 snapshots to compare.")

            if snap_b_date:
                cache_key_a = f"snap_data_{snap_a_date}"
                cache_key_b = f"snap_data_{snap_b_date}"
                if cache_key_a not in st.session_state:
                    with st.spinner(f"Loading {_fmt_snap(snap_a_date)}..."):
                        st.session_state[cache_key_a] = load_pipeline_snapshot(snap_a_date)
                if cache_key_b not in st.session_state:
                    with st.spinner(f"Loading {_fmt_snap(snap_b_date)}..."):
                        st.session_state[cache_key_b] = load_pipeline_snapshot(snap_b_date)

                try:
                    snap_a = st.session_state[cache_key_a]
                    snap_b = st.session_state[cache_key_b]

                    def _snap_totals(snap, keys, months):
                        frames = [snap[k].copy() for k in keys if k in snap and isinstance(snap[k], pd.DataFrame)]
                        if not frames:
                            return pd.Series(dtype=float)
                        combined = pd.concat(frames, ignore_index=True)
                        for m in MONTHS:
                            if m in combined.columns:
                                combined[m] = pd.to_numeric(combined[m], errors="coerce").fillna(0)
                        cols = [m for m in months if m in combined.columns]
                        return combined.groupby("Client")[cols].sum().sum(axis=1)

                    totals_a = _snap_totals(snap_a, rv_keys, period_months)
                    totals_b = _snap_totals(snap_b, rv_keys, period_months)
                    cmp_df, cmp_sel = _render_comparison(
                        totals_a, totals_b,
                        _fmt_snap(snap_a_date), _fmt_snap(snap_b_date),
                        table_key="cmp_snap_table",
                    )

                    # ── Project drill-down ────────────────────────────────
                    if cmp_sel and cmp_sel.get("selection", {}).get("rows"):
                        sel_client = cmp_df.iloc[cmp_sel["selection"]["rows"][0]]["Client"]

                        def _proj_totals(snap, keys, client, months):
                            frames = [snap[k].copy() for k in keys if k in snap and isinstance(snap[k], pd.DataFrame)]
                            if not frames:
                                return pd.Series(dtype=float, name=client)
                            combined = pd.concat(frames, ignore_index=True)
                            combined = combined[combined["Client"] == client]
                            for m in MONTHS:
                                if m in combined.columns:
                                    combined[m] = pd.to_numeric(combined[m], errors="coerce").fillna(0)
                            cols = [m for m in months if m in combined.columns]
                            if "Project" not in combined.columns or combined.empty:
                                return pd.Series(dtype=float)
                            return combined.groupby("Project")[cols].sum().sum(axis=1)

                        proj_a = _proj_totals(snap_a, rv_keys, sel_client, period_months)
                        proj_b = _proj_totals(snap_b, rv_keys, sel_client, period_months)
                        all_projs = sorted(set(proj_a.index) | set(proj_b.index))

                        col_a_lbl = _fmt_snap(snap_a_date)
                        col_b_lbl = _fmt_snap(snap_b_date)
                        proj_rows = []
                        for proj in all_projs:
                            va = proj_a.get(proj, 0)
                            vb = proj_b.get(proj, 0)
                            ch = va - vb
                            proj_rows.append({
                                "Project": proj,
                                col_a_lbl: fmt_gbp(va) if va else "—",
                                col_b_lbl: fmt_gbp(vb) if vb else "—",
                                "Change": (("+" if ch > 0 else "") + fmt_gbp(ch)) if ch != 0 else "—",
                                "_ch": ch,
                                "_new": vb == 0 and va > 0,
                                "_lost": va == 0 and vb > 0,
                            })
                        proj_df = pd.DataFrame(proj_rows).sort_values("_ch", ascending=False)

                        def _style_proj(row):
                            ch = proj_df.loc[proj_df["Project"] == row["Project"], "_ch"]
                            cv = ch.iloc[0] if not ch.empty else 0
                            return ["color:#27ae60;font-weight:bold" if cv > 0 else ("color:#c0392b;font-weight:bold" if cv < 0 else "") if c == "Change" else "" for c in row.index]

                        with st.expander(f"**{sel_client}** — project breakdown", expanded=True):
                            new_p  = proj_df[proj_df["_new"]]["Project"].tolist()
                            lost_p = proj_df[proj_df["_lost"]]["Project"].tolist()
                            if new_p:
                                st.success(f"New projects: {', '.join(new_p)}")
                            if lost_p:
                                st.warning(f"Dropped projects: {', '.join(lost_p)}")
                            st.dataframe(
                                proj_df[["Project", col_a_lbl, col_b_lbl, "Change"]].style.apply(_style_proj, axis=1),
                                use_container_width=True, hide_index=True,
                            )

                except Exception as e:
                    st.error(f"Could not load snapshots: {e}")

        # ── Project drill-down ────────────────────────────────────────────────
        st.divider()
        st.caption("💡 Click a client row to see their project breakdown")

        if sel_rv and sel_rv.get("selection", {}).get("rows"):
            row_idx_rv = sel_rv["selection"]["rows"][0]
            sel_client = rv_by_client.iloc[row_idx_rv]["Client"]
            sp_token   = st.session_state.get("sp_token")

            with st.expander(f"**{sel_client}** — Project Breakdown", expanded=True):
                sel_drill_month = st.selectbox(
                    "Month", MONTHS,
                    index=MONTHS.index(cur_month_name),
                    key="rv_drill_month",
                )

                if sp_token:
                    with st.spinner("Loading..."):
                        try:
                            projects = fetch_client_projects(sp_token, sel_client, sel_drill_month)
                            if not projects:
                                st.info(f"No project data found for {sel_client} in {sel_drill_month}.")
                            else:
                                all_proj_rows = []
                                for proj_list in projects.values():
                                    for p in proj_list:
                                        all_proj_rows.append({"Project": p["Project"], "Revenue": p["Revenue"]})
                                proj_df = pd.DataFrame(all_proj_rows).groupby("Project", as_index=False)["Revenue"].sum()
                                proj_df = proj_df.sort_values("Revenue", ascending=False)
                                total_proj = proj_df["Revenue"].sum()
                                proj_df["Revenue"] = proj_df["Revenue"].apply(lambda v: f"£{v:,.0f}")
                                st.dataframe(proj_df, use_container_width=True, hide_index=True)
                                st.markdown(f"**Total: £{total_proj:,.0f}**")
                        except Exception as e:
                            st.error(f"Could not load project breakdown: {e}")
                else:
                    st.info("Authenticate via SharePoint to see project-level breakdown.")

# ── Capacity Planning tab ─────────────────────────────────────────────────────
with tab_sow:
    st.subheader("Capacity Planning")
    st.caption("Upload Scope of Work files to see hours and fees by department and role across projects.")

    # ── Neverland SoW template parser ────────────────────────────────────────
    def _parse_neverland_sow(file_bytes: bytes, filename: str):
        """
        Parses the standard Neverland 'Staffing Fee Template' Excel format.
        Returns a DataFrame with columns: Project, Client, Department, Role,
        Hourly Rate, Total Hours, Total Fee.
        Column positions (0-indexed): dept=0, role=2, rate=4, hours=30, fee=32.
        """
        import io, re
        try:
            xf = pd.ExcelFile(io.BytesIO(file_bytes))
        except Exception:
            return None

        # ── extract project + client from Overview sheet ──────────────────
        project_name = filename.rsplit(".", 1)[0]
        client_name  = ""
        if "Overview" in xf.sheet_names:
            try:
                ov = xf.parse("Overview", header=None)
                for _, row in ov.iterrows():
                    vals = [str(v).strip() for v in row if v is not None and str(v).strip()]
                    joined = " ".join(vals)
                    if re.search(r"^client", joined, re.I) and len(vals) >= 2:
                        client_name = vals[-1]
                    if re.search(r"^project", joined, re.I) and len(vals) >= 2:
                        raw_proj = vals[-1]
                        project_name = re.sub(r"^project:\s*", "", raw_proj, flags=re.I).strip()
            except Exception:
                pass

        # ── parse Staffing Fee Template sheet ────────────────────────────
        if "Staffing Fee Template" not in xf.sheet_names:
            return None
        try:
            ws_raw = xf.parse("Staffing Fee Template", header=None)
        except Exception:
            return None

        rows_out = []
        current_dept = ""
        DEPT_COL, ROLE_COL, RATE_COL = 0, 2, 4

        # Auto-detect total Hours and Fee columns from the header row.
        # The header row has "Hourly Rate" in col 4. After that, find the
        # LAST "Hours" column (before any FTE column) and the Fee column after it.
        HRS_COL, FEE_COL = 30, 32  # sensible defaults
        for _, hrow in ws_raw.iterrows():
            vals = [str(v).strip().lower() if v is not None else "" for v in hrow]
            if "hourly rate" in vals or (len(vals) > 4 and vals[4] in ("hourly rate",)):
                # Find last 'hours' column
                last_hrs = None
                for ci, v in enumerate(vals):
                    if v == "hours":
                        last_hrs = ci
                if last_hrs is not None:
                    HRS_COL = last_hrs
                    # Fee column is the next non-empty, non-FTE column after last_hrs
                    for ci in range(last_hrs + 1, len(vals)):
                        v = vals[ci]
                        if v and "fte" not in v and "hours" not in v:
                            FEE_COL = ci
                            break
                break

        for _, row in ws_raw.iterrows():
            dept_val = row.iloc[DEPT_COL] if DEPT_COL < len(row) else None
            role_val = row.iloc[ROLE_COL] if ROLE_COL < len(row) else None
            rate_val = row.iloc[RATE_COL] if RATE_COL < len(row) else None
            hrs_val  = row.iloc[HRS_COL]  if HRS_COL  < len(row) else None
            fee_val  = row.iloc[FEE_COL]  if FEE_COL  < len(row) else None

            role_str = str(role_val).strip() if role_val is not None and str(role_val).strip() != "nan" else ""
            dept_str = str(dept_val).strip() if dept_val is not None and str(dept_val).strip() != "nan" else ""

            # Section header row: has a role/section label but no hourly rate
            if role_str and (rate_val is None or str(rate_val).strip() == "nan"):
                if role_str.upper() != "TOTAL" and role_str not in ("Department / Job Title", "NAME"):
                    current_dept = role_str
                continue

            # Skip header rows and TOTAL subtotal rows
            if not role_str or role_str.upper() == "TOTAL":
                continue

            try:
                rate = float(rate_val) if rate_val is not None else 0.0
                hrs  = float(hrs_val)  if hrs_val  is not None else 0.0
                fee  = float(fee_val)  if fee_val  is not None else 0.0
            except (TypeError, ValueError):
                continue

            if hrs <= 0:
                continue

            dept = dept_str if dept_str else current_dept
            _DEPT_RENAME = {"Account": "Client Service", "Total Account": "Client Service"}
            dept = _DEPT_RENAME.get(dept, dept)
            # Roles with "Strategy" in their title belong in Strategy, not Client Service
            if dept == "Client Service" and "strategy" in role_str.lower():
                dept = "Strategy"
            rows_out.append({
                "Project":      project_name,
                "Client":       client_name,
                "Department":   dept,
                "Role":         role_str,
                "Hourly Rate":  rate,
                "Total Hours":  hrs,
                "Total Fee":    fee,
            })

        if not rows_out:
            return None
        return pd.DataFrame(rows_out)

    # ── generic fallback parser ───────────────────────────────────────────────
    def _parse_generic_sow(file_bytes: bytes, filename: str):
        import io
        _alias = lambda cols, names: next(
            (c for c in cols if c.strip().lower().replace(" ", "_").replace("-", "_") in names), None
        )
        ROLE_A  = {"role","job_title","job_role","position","resource","discipline","function"}
        HOURS_A = {"hours","total_hours","estimated_hours","hrs","planned_hours","allocated_hours"}
        try:
            if filename.lower().endswith(".csv"):
                df = pd.read_csv(io.BytesIO(file_bytes))
            else:
                xf = pd.ExcelFile(io.BytesIO(file_bytes))
                df = None
                for sheet in xf.sheet_names:
                    c = xf.parse(sheet)
                    if _alias(c.columns, ROLE_A) and _alias(c.columns, HOURS_A):
                        df = c; break
                if df is None:
                    df = xf.parse(xf.sheet_names[0])
        except Exception:
            return None

        role_col  = _alias(df.columns, ROLE_A)
        hours_col = _alias(df.columns, HOURS_A)
        if not role_col or not hours_col:
            return None

        out = df[[role_col, hours_col]].copy()
        out.columns = ["Role", "Total Hours"]
        out.insert(0, "Project",    filename.rsplit(".", 1)[0])
        out.insert(1, "Client",     "")
        out.insert(2, "Department", "")
        out["Hourly Rate"] = 0.0
        out["Total Fee"]   = 0.0
        out["Total Hours"] = pd.to_numeric(out["Total Hours"], errors="coerce").fillna(0)
        out = out[out["Total Hours"] > 0].dropna(subset=["Role"])
        return out[["Project","Client","Department","Role","Hourly Rate","Total Hours","Total Fee"]]

    def _parse_uploaded_sow(uploaded):
        name = uploaded.name
        data = uploaded.read()
        ext  = name.rsplit(".", 1)[-1].lower()
        if ext in ("xlsx", "xls"):
            df = _parse_neverland_sow(data, name)
            if df is None or df.empty:
                df = _parse_generic_sow(data, name)
        elif ext == "csv":
            df = _parse_generic_sow(data, name)
        else:
            df = None
        if df is None or df.empty:
            return None, name
        return df, None

    # ── session state — load from GitHub on first visit ───────────────────────
    SOW_COLS = ["Project","Client","Department","Role","Hourly Rate","Total Hours","Total Fee","_source"]
    if "sow_data" not in st.session_state:
        try:
            _persisted = load_sow_data()
            st.session_state["sow_data"] = _persisted if not _persisted.empty else pd.DataFrame(columns=SOW_COLS)
        except Exception:
            st.session_state["sow_data"] = pd.DataFrame(columns=SOW_COLS)

    # ── file uploader ─────────────────────────────────────────────────────────
    uploaded_files = st.file_uploader(
        "Upload Scope of Work files",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        help="Use your standard Neverland Staffing Fee Template Excel file.",
    )

    if uploaded_files:
        added = False
        for uf in uploaded_files:
            if uf.name in st.session_state["sow_data"]["_source"].values:
                continue
            df_parsed, failed = _parse_uploaded_sow(uf)
            if failed:
                st.warning(f"Could not extract data from **{failed}**. Make sure it uses the Neverland Staffing Fee Template format.")
            else:
                df_parsed["_source"] = uf.name
                st.session_state["sow_data"] = pd.concat(
                    [st.session_state["sow_data"], df_parsed], ignore_index=True
                )
                n   = len(df_parsed)
                hrs = df_parsed["Total Hours"].sum()
                fee = df_parsed["Total Fee"].sum()
                st.success(f"Loaded **{uf.name}** — {n} roles, {hrs:,.0f} hours, £{fee:,.0f} total fee")
                added = True
        if added:
            try:
                save_sow_data(st.session_state["sow_data"])
            except Exception as _e:
                st.warning(f"Could not save SoW data: {_e}")

    sow_df = st.session_state["sow_data"].copy()
    # ensure numeric
    for _nc in ("Hourly Rate", "Total Hours", "Total Fee"):
        sow_df[_nc] = pd.to_numeric(sow_df[_nc], errors="coerce").fillna(0)

    if sow_df.empty:
        st.info("No data yet. Upload a Scope of Work file above to get started.")
    else:
        # ── manage loaded files ───────────────────────────────────────────────
        sources = sow_df["_source"].unique().tolist()
        with st.expander(f"Manage loaded files ({len(sources)})", expanded=False):
            for src in sources:
                ca, cb = st.columns([6, 1])
                sub = sow_df[sow_df["_source"] == src]
                ca.markdown(f"**{sub['Project'].iloc[0]}** ({src})")
                if cb.button("Remove", key=f"rm_sow_{src}"):
                    st.session_state["sow_data"] = st.session_state["sow_data"][
                        st.session_state["sow_data"]["_source"] != src
                    ].reset_index(drop=True)
                    try:
                        save_sow_data(st.session_state["sow_data"])
                    except Exception:
                        pass
                    st.rerun()

        @st.dialog("Role Breakdown by Client", width="large")
        def _show_role_breakdown(role_name, dept_name):
            st.subheader(f"{role_name} — {dept_name}")
            role_rows = sow_df[(sow_df["Role"] == role_name) & (sow_df["Department"] == dept_name)].copy()
            for col in ("Total Hours", "Total Fee"):
                role_rows[col] = pd.to_numeric(role_rows[col], errors="coerce").fillna(0)
            breakdown = (
                role_rows.groupby(["Project", "Client"] if "Client" in role_rows.columns else ["Project"])
                .agg(Hours=("Total Hours", "sum"), Fee=("Total Fee", "sum"))
                .reset_index()
                .sort_values("Fee", ascending=False)
            )
            total_h = breakdown["Hours"].sum()
            total_f = breakdown["Fee"].sum()
            breakdown["Hours"] = breakdown["Hours"].apply(lambda v: f"{v:,.0f}")
            breakdown["Fee"]   = breakdown["Fee"].apply(lambda v: f"£{v:,.0f}")
            st.dataframe(breakdown, use_container_width=True, hide_index=True)
            st.markdown(f"**Total — {total_h:,.0f} hrs · £{total_f:,.0f}**")

        # ── department sections ───────────────────────────────────────────────
        depts = sorted(d for d in sow_df["Department"].unique() if d)
        for dept in depts:
            dept_df = sow_df[sow_df["Department"] == dept]
            dept_h = dept_df["Total Hours"].sum()
            dept_f = dept_df["Total Fee"].sum()

            st.markdown(
                f"<div style='background:{NV_DARK};color:#fff;padding:8px 14px;"
                f"border-left:4px solid {NV_PINK};border-radius:4px;margin-top:20px'>"
                f"<b>{dept}</b> &nbsp;·&nbsp; {dept_h:,.0f} hrs &nbsp;·&nbsp; £{dept_f:,.0f}"
                f"</div>",
                unsafe_allow_html=True,
            )

            role_df_raw = (
                dept_df.groupby("Role")
                .agg(Hours=("Total Hours", "sum"), Fee=("Total Fee", "sum"))
                .reset_index()
                .sort_values("Fee", ascending=False)
            )
            display_role = role_df_raw.copy()
            display_role["Hours"] = display_role["Hours"].apply(lambda v: f"{v:,.0f}")
            display_role["Fee"]   = display_role["Fee"].apply(lambda v: f"£{v:,.0f}")
            st.caption("Click a row to see the client breakdown.")
            sel = st.dataframe(
                display_role, use_container_width=True, hide_index=True,
                on_select="rerun", selection_mode="single-row",
                key=f"sow_role_{dept}",
            )
            if sel and sel.get("selection", {}).get("rows"):
                selected_role = role_df_raw.iloc[sel["selection"]["rows"][0]]["Role"]
                _show_role_breakdown(selected_role, dept)

        # ── grand total ───────────────────────────────────────────────────────
        st.markdown("---")
        gt_h = sow_df["Total Hours"].sum()
        gt_f = sow_df["Total Fee"].sum()
        st.markdown(f"**Grand Total — {gt_h:,.0f} hours | £{gt_f:,.0f}**")


# ── Floating AI Chat ──────────────────────────────────────────────────────────
def _build_chat_context():
    parts = []
    if "pipeline" in st.session_state:
        for _status, _df in st.session_state["pipeline"].items():
            if not isinstance(_df, pd.DataFrame) or _df.empty:
                continue
            _month_cols = [m for m in MONTHS if m in _df.columns]
            if _month_cols:
                _num = _df[_month_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
                _annual = _num.values.sum()
                parts.append(f"\nPipeline status: {_status} — Annual total: £{_annual:,.0f}")
                # Monthly totals
                _monthly = _num.sum()
                for _m in MONTHS:
                    if _m in _monthly.index:
                        parts.append(f"  {_m}: £{_monthly[_m]:,.0f}")
                # Per-client breakdown
                if "Client" in _df.columns:
                    for _, _row in _df.iterrows():
                        _client = _row.get("Client", "Unknown")
                        _project = _row.get("Project", "")
                        _label = f"{_client}" + (f" / {_project}" if _project else "")
                        _row_nums = pd.to_numeric(_row[_month_cols], errors="coerce").fillna(0)
                        _row_total = _row_nums.sum()
                        if _row_total > 0:
                            _mvals = ", ".join(f"{m}:£{_row_nums[m]:,.0f}" for m in MONTHS if m in _row_nums.index and _row_nums[m] > 0)
                            parts.append(f"  {_label}: £{_row_total:,.0f} total ({_mvals})")
    if "sow_data" in st.session_state and not st.session_state["sow_data"].empty:
        _sow = st.session_state["sow_data"]
        _hrs = pd.to_numeric(_sow["Total Hours"], errors="coerce").sum()
        _fee = pd.to_numeric(_sow["Total Fee"], errors="coerce").sum()
        _projs = _sow["Project"].nunique() if "Project" in _sow.columns else 0
        parts.append(f"Scope of Works: {_projs} projects, {_hrs:,.0f} total hours, £{_fee:,.0f} total fee")
        for _d, _row in _sow.groupby("Department").agg(Hours=("Total Hours","sum"), Fee=("Total Fee","sum")).iterrows():
            if _d:
                parts.append(f"  Dept {_d}: {pd.to_numeric(_row['Hours'], errors='coerce'):,.0f} hrs, £{pd.to_numeric(_row['Fee'], errors='coerce'):,.0f}")
    if "budget" in st.session_state and isinstance(st.session_state["budget"], pd.DataFrame):
        parts.append(f"Budget data: {len(st.session_state['budget'])} rows loaded")
    if "capacity" in st.session_state and isinstance(st.session_state["capacity"], pd.DataFrame):
        parts.append(f"Capacity data: {len(st.session_state['capacity'])} rows loaded")
    if "salary_dept" in st.session_state and isinstance(st.session_state["salary_dept"], pd.DataFrame):
        parts.append(f"Salary by dept: {len(st.session_state['salary_dept'])} rows loaded")
    return "\n".join(parts) if parts else "No dashboard data is currently loaded."


@st.dialog("Neverland AI", width="small")
def _chat_popup():
    if "chat_messages" not in st.session_state:
        st.session_state["chat_messages"] = []

    _chat_area = st.container(height=400)
    with _chat_area:
        for _msg in st.session_state["chat_messages"]:
            with st.chat_message(_msg["role"]):
                st.markdown(_msg["content"])

    _user_input = st.chat_input("Ask about your data...")
    if st.button("Clear", key="clear_chat_popup"):
        st.session_state["chat_messages"] = []
        st.rerun()

    if _user_input:
        st.session_state["chat_messages"].append({"role": "user", "content": _user_input})
        _system_prompt = (
            "You are the Neverland Finance AI assistant, embedded in Neverland Creative Agency's "
            "finance dashboard. You help the team understand their financial data — pipeline revenue, "
            "capacity, budgets, and scope of works. Be concise and direct. Format numbers clearly.\n\n"
            "Current dashboard data:\n" + _build_chat_context()
        )
        try:
            import anthropic as _anthropic
            _api_key = str(st.secrets.get("ANTHROPIC_API_KEY", "")).strip()
            _client = _anthropic.Anthropic(api_key=_api_key)
            _history = [{"role": m["role"], "content": m["content"]} for m in st.session_state["chat_messages"]]
            with _chat_area:
                with st.chat_message("user"):
                    st.markdown(_user_input)
                with st.chat_message("assistant"):
                    _placeholder = st.empty()
                    _full = ""
                    with _client.messages.stream(
                        model="claude-opus-5",
                        max_tokens=1024,
                        system=_system_prompt,
                        messages=_history,
                    ) as _stream:
                        for _chunk in _stream.text_stream:
                            _full += _chunk
                            _placeholder.markdown(_full + "▌")
                    _placeholder.markdown(_full)
            st.session_state["chat_messages"].append({"role": "assistant", "content": _full})
        except Exception as _e:
            st.error(f"AI error: {_e}")


if st.button("💬", key="open_chat_fab"):
    _chat_popup()

# Reposition the FAB button to bottom-right corner via JS
import streamlit.components.v1 as _components
_components.html("""
<script>
(function() {
  function applyFab() {
    var doc = window.parent.document;
    var btns = doc.querySelectorAll('button');
    for (var i = 0; i < btns.length; i++) {
      if (btns[i].textContent.trim() === '💬') {
        var el = btns[i].closest('[data-testid="element-container"]') || btns[i];
        el.style.setProperty('position', 'fixed', 'important');
        el.style.setProperty('bottom', '28px', 'important');
        el.style.setProperty('right', '28px', 'important');
        el.style.setProperty('z-index', '9999', 'important');
        el.style.setProperty('width', 'auto', 'important');
        btns[i].style.setProperty('border-radius', '50%', 'important');
        btns[i].style.setProperty('width', '56px', 'important');
        btns[i].style.setProperty('height', '56px', 'important');
        btns[i].style.setProperty('min-height', 'unset', 'important');
        btns[i].style.setProperty('background', '#E8147B', 'important');
        btns[i].style.setProperty('color', 'white', 'important');
        btns[i].style.setProperty('font-size', '22px', 'important');
        btns[i].style.setProperty('padding', '0', 'important');
        btns[i].style.setProperty('border', 'none', 'important');
        btns[i].style.setProperty('box-shadow', '0 4px 16px rgba(0,0,0,0.25)', 'important');
        btns[i].style.setProperty('cursor', 'pointer', 'important');
        return;
      }
    }
    setTimeout(applyFab, 200);
  }
  applyFab();
})();
</script>
""", height=0)
