"""
Infinity Drones — Personal Monthly Budget Dashboard
Mobile-friendly Streamlit app for net income, expenses, savings, checkpoints, and trends.
"""

import os
import sqlite3
from datetime import date, datetime
from dateutil.relativedelta import relativedelta

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

DB_PATH = os.path.join(os.path.dirname(__file__), "budget.db")

st.set_page_config(
    page_title="Infinity Drones Budget",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── Styling: mobile-first, portrait + landscape friendly ─────────────────────
st.markdown(
    """
    <style>
    :root {
        --card-bg: #ffffff;
        --card-border: #e5e7eb;
        --muted: #6b7280;
        --green: #16a34a;
        --red: #dc2626;
        --blue: #2563eb;
        --amber: #d97706;
    }
    .block-container {
        padding-top: 1.2rem;
        padding-left: clamp(0.75rem, 2vw, 3rem);
        padding-right: clamp(0.75rem, 2vw, 3rem);
        max-width: 1400px;
    }
    h1, h2, h3 { letter-spacing: -0.02em; }
    .hero {
        padding: 1rem 1.1rem;
        border: 1px solid var(--card-border);
        border-radius: 18px;
        background: linear-gradient(135deg, #f8fafc 0%, #eef6ff 100%);
        margin-bottom: 1rem;
    }
    .hero-title { font-size: clamp(1.55rem, 5vw, 2.4rem); font-weight: 800; margin-bottom: .25rem; }
    .hero-subtitle { color: var(--muted); font-size: clamp(.9rem, 2.5vw, 1rem); }
    .metric-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(155px, 1fr));
        gap: .75rem;
        margin: .75rem 0 1rem 0;
    }
    .metric-card {
        border: 1px solid var(--card-border);
        border-radius: 16px;
        padding: .95rem;
        background: var(--card-bg);
        box-shadow: 0 1px 4px rgba(15,23,42,.06);
        min-height: 105px;
    }
    .metric-label { color: var(--muted); font-size: .82rem; font-weight: 650; text-transform: uppercase; letter-spacing: .04em; }
    .metric-value { font-size: clamp(1.35rem, 6vw, 2rem); font-weight: 800; line-height: 1.1; margin-top: .35rem; word-break: keep-all; }
    .metric-note { color: var(--muted); font-size: .78rem; margin-top: .4rem; }
    .good { color: var(--green); } .bad { color: var(--red); } .blue { color: var(--blue); } .amber { color: var(--amber); }
    .small-note {
        padding: .75rem .9rem;
        border-left: 4px solid #2563eb;
        border-radius: 10px;
        background: #eff6ff;
        color: #1e3a8a;
        margin: .65rem 0 1rem 0;
        font-size: .92rem;
    }
    .warning-note {
        padding: .75rem .9rem;
        border-left: 4px solid #d97706;
        border-radius: 10px;
        background: #fffbeb;
        color: #78350f;
        margin: .65rem 0 1rem 0;
        font-size: .92rem;
    }
    div[data-testid="stDataFrame"] { width: 100%; overflow-x: auto; }
    @media (max-width: 720px) {
        .block-container { padding-top: .7rem; }
        .hero { padding: .85rem; border-radius: 14px; }
        .metric-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: .55rem; }
        .metric-card { padding: .75rem; min-height: 95px; }
        .metric-label { font-size: .72rem; }
        .metric-note { font-size: .72rem; }
    }
    @media (max-width: 420px) {
        .metric-grid { grid-template-columns: 1fr; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ─── DB Helpers ───────────────────────────────────────────────────────────────
def get_conn():
    return sqlite3.connect(DB_PATH)


def init_db():
    conn = get_conn()
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS budget_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            month TEXT NOT NULL,
            category TEXT NOT NULL,
            subcategory TEXT DEFAULT '',
            amount REAL NOT NULL,
            entry_type TEXT NOT NULL,
            notes TEXT DEFAULT '',
            source TEXT DEFAULT 'manual',
            is_sample INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS monthly_snapshots (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            month TEXT NOT NULL UNIQUE,
            total_income REAL DEFAULT 0,
            total_expenses REAL DEFAULT 0,
            total_savings REAL DEFAULT 0,
            net_cashflow REAL DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        """
    )
    conn.commit()
    conn.close()


init_db()


def add_entry(month, category, subcategory, amount, entry_type, notes="", source="manual", is_sample=0):
    conn = get_conn()
    conn.execute(
        """
        INSERT INTO budget_entries
            (month, category, subcategory, amount, entry_type, notes, source, is_sample)
        VALUES (?,?,?,?,?,?,?,?)
        """,
        (month, category, subcategory, amount, entry_type, notes, source, is_sample),
    )
    conn.commit()
    conn.close()


def delete_entry(entry_id):
    conn = get_conn()
    conn.execute("DELETE FROM budget_entries WHERE id = ?", (entry_id,))
    conn.commit()
    conn.close()


def get_entries(month=None, entry_type=None):
    conn = get_conn()
    query = "SELECT * FROM budget_entries WHERE 1=1"
    params = []
    if month:
        query += " AND month = ?"
        params.append(month)
    if entry_type:
        query += " AND entry_type = ?"
        params.append(entry_type)
    query += " ORDER BY entry_type, category, subcategory"
    df = pd.read_sql(query, conn, params=params)
    conn.close()
    return df


def get_all_entries():
    conn = get_conn()
    df = pd.read_sql("SELECT * FROM budget_entries ORDER BY month DESC, entry_type, category", conn)
    conn.close()
    return df


def get_months() -> list[str]:
    conn = get_conn()
    df = pd.read_sql("SELECT DISTINCT month FROM budget_entries ORDER BY month", conn)
    conn.close()
    # str() guard: legacy imported months must stay strings for sorting/reversal.
    return [str(m) for m in df["month"].tolist()]


def rebuild_monthly_checkpoints():
    """Make checkpoint values explicit and auditable for every month."""
    conn = get_conn()
    rows = conn.execute(
        """
        SELECT month,
               SUM(CASE WHEN entry_type='income' THEN amount ELSE 0 END) AS income,
               SUM(CASE WHEN entry_type='expense' THEN amount ELSE 0 END) AS expenses,
               SUM(CASE WHEN entry_type='savings' THEN amount ELSE 0 END) AS savings
        FROM budget_entries
        GROUP BY month
        ORDER BY month
        """
    ).fetchall()
    for month, income, expenses, savings in rows:
        income = income or 0
        expenses = expenses or 0
        savings = savings or 0
        net_cashflow = income - expenses - savings
        conn.execute(
            """
            INSERT INTO monthly_snapshots (month, total_income, total_expenses, total_savings, net_cashflow)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(month) DO UPDATE SET
                total_income=excluded.total_income,
                total_expenses=excluded.total_expenses,
                total_savings=excluded.total_savings,
                net_cashflow=excluded.net_cashflow
            """,
            (month, income, expenses, savings, net_cashflow),
        )
    conn.commit()
    conn.close()


def get_checkpoint_df():
    rebuild_monthly_checkpoints()
    conn = get_conn()
    df = pd.read_sql(
        """
        SELECT month,
               total_income AS income,
               total_expenses AS expenses,
               total_savings AS savings,
               net_cashflow AS net
        FROM monthly_snapshots
        ORDER BY month
        """,
        conn,
    )
    conn.close()
    if not df.empty:
        df["cash_available"] = df["income"] - df["expenses"]
        df["savings_rate"] = (df["savings"] / df["income"] * 100).where(df["income"] > 0, 0).round(1)
        df["expense_ratio"] = (df["expenses"] / df["income"] * 100).where(df["income"] > 0, 0).round(1)
    return df


# ─── Helpers ──────────────────────────────────────────────────────────────────
def zar(value):
    return f"R {float(value):,.0f}"


def pct(value):
    return f"{float(value):.1f}%"


def render_metric_grid(metrics):
    html = ['<div class="metric-grid">']
    for label, value, note, cls in metrics:
        # Keep HTML left-aligned. Indented multi-line HTML can be interpreted by
        # Streamlit/Markdown as a code block after the first rendered element.
        html.append(
            f'<div class="metric-card">'
            f'<div class="metric-label">{label}</div>'
            f'<div class="metric-value {cls}">{value}</div>'
            f'<div class="metric-note">{note}</div>'
            f'</div>'
        )
    html.append("</div>")
    st.markdown("".join(html), unsafe_allow_html=True)


def responsive_chart_layout(fig, height=360, show_legend=True):
    fig.update_layout(
        height=height,
        autosize=True,
        margin=dict(l=20, r=16, t=45, b=45),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0) if show_legend else None,
        hovermode="x unified",
        font=dict(size=12),
    )
    fig.update_xaxes(tickangle=-30, automargin=True)
    fig.update_yaxes(tickprefix="R ", separatethousands=True, automargin=True)
    return fig


def value_text(series):
    return [zar(v) for v in series]


def money_column_config(*column_names):
    """Streamlit dataframe formatting for Rand values: currency, max 2 decimals."""
    return {
        name: st.column_config.NumberColumn(name, format="R %.2f")
        for name in column_names
    }


DEFAULT_INCOME = ["Cheque Credit / Net Salary", "Side Business", "Investment Income", "Other Income"]
DEFAULT_FIXED = [
    "Rent / Mortgage", "Insurance", "Subscriptions", "Loan Payments", "School Fees",
    "Medical Aid", "Rates & Taxes", "Internet & Phone"
]
DEFAULT_VARIABLE = [
    "Groceries", "Fuel / Transport", "Electricity & Water", "Entertainment", "Kids Activities",
    "Clothing", "Medical (Out of Pocket)", "Personal Care", "Gifts", "Miscellaneous"
]
DEFAULT_SAVINGS = ["Emergency Fund", "Investments", "Retirement", "Other Savings"]

# ─── Sidebar ──────────────────────────────────────────────────────────────────
st.sidebar.title("💰 Infinity Drones")
st.sidebar.subheader("Personal Budget")
page = st.sidebar.radio(
    "Navigate",
    ["📊 Dashboard", "➕ Add Entries", "📈 Trends & Checkpoints", "📋 Data Table", "⚙️ Settings"],
)

# ─── Dashboard ────────────────────────────────────────────────────────────────
if page == "📊 Dashboard":
    st.markdown(
        """
        <div class="hero">
            <div class="hero-title">📊 Personal Budget Dashboard</div>
            <div class="hero-subtitle">Uses net money received — cheque credit / actual salary paid into the account — not gross salary before tax and deductions.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    months = get_months()
    if not months:
        st.info("Start by adding entries under **➕ Add Entries**.")
        st.stop()

    all_entries = get_all_entries()
    salary_months = sorted(
        all_entries.loc[
            (all_entries["entry_type"] == "income")
            & (
                all_entries["category"].str.contains("cheque|salary", case=False, na=False)
                | all_entries["subcategory"].str.contains("cheque|salary", case=False, na=False)
                | all_entries["notes"].str.contains("cheque|salary", case=False, na=False)
            ),
            "month",
        ].unique().tolist()
    )
    # Newest-first for the picker only; get_months() stays ascending for charts.
    view_options = ["Average overview"] + months[::-1]
    selected_view = st.selectbox(
        "View",
        view_options,
        index=0,
        help="Open on the average picture first. Pick a month when you want the detailed month view.",
    )
    is_average_view = selected_view == "Average overview"
    if is_average_view:
        analysis_months = salary_months or months
        df = all_entries[all_entries["month"].isin(analysis_months)].copy()
        period_divisor = max(len(analysis_months), 1)
        period_label = f"Average across {period_divisor} salary months"
    else:
        analysis_months = [selected_view]
        df = get_entries(selected_view)
        period_divisor = 1
        period_label = selected_view
    if df.empty:
        st.warning(f"No entries for {selected_view}.")
        st.stop()

    sample_count = int(df["is_sample"].sum()) if "is_sample" in df.columns else 0
    real_count = int((df["source"] == "user_file").sum()) if "source" in df.columns else 0
    if real_count and not sample_count:
        st.markdown(
            """
            <div class="small-note">
            <b>Real data imported from the Infinity Drones Excel workbook.</b>
            Salary comes from actual cheque-credit salary deposits only; business income / other bank credits are ignored. Business-coded categories such as stock, courier, DJI/FPV rows and SARS were excluded from expenses. The R1,400 scheduled monthly transfer is shown as savings.
            </div>
            """,
            unsafe_allow_html=True,
        )
    elif real_count and sample_count:
        st.markdown(
            """
            <div class="warning-note">
            <b>Real rows are present, but seeded sample rows are still visible for this month.</b>
            Delete or ignore sample rows before using the dashboard for decisions.
            </div>
            """,
            unsafe_allow_html=True,
        )
    if sample_count:
        st.markdown(
            """
            <div class="warning-note">
            <b>Current data is sample/seeded data, not actual financial records.</b>
            The visible numbers were generated by me as placeholders. Do not use them for decisions until you send the real bank/payslip/budget records.
            </div>
            """,
            unsafe_allow_html=True,
        )

    income_df = df[df["entry_type"] == "income"]
    salary_mask = (
        income_df["category"].str.contains("cheque|salary", case=False, na=False)
        | income_df["subcategory"].str.contains("cheque|salary", case=False, na=False)
        | income_df["notes"].str.contains("cheque|salary", case=False, na=False)
    )
    salary_total_raw = income_df.loc[salary_mask, "amount"].sum()
    other_income_total_raw = income_df.loc[~salary_mask, "amount"].sum()
    expense_total_raw = df.loc[df["entry_type"] == "expense", "amount"].sum()
    savings_total_raw = df.loc[df["entry_type"] == "savings", "amount"].sum()

    salary_total = salary_total_raw / period_divisor
    other_income_total = other_income_total_raw / period_divisor
    income_total = salary_total + other_income_total
    expense_total = expense_total_raw / period_divisor
    savings_total = savings_total_raw / period_divisor
    cash_available = income_total - expense_total
    net = cash_available - savings_total
    savings_rate = (savings_total / income_total * 100) if income_total > 0 else 0
    expense_ratio = (expense_total / income_total * 100) if income_total > 0 else 0

    render_metric_grid([
        ("Avg net salary / cheque credit" if is_average_view else "Net salary / cheque credit", zar(salary_total), f"{period_label}; actual salary paid into bank account", "good"),
        ("Avg other income" if is_average_view else "Other income", zar(other_income_total), "Business income excluded", "blue"),
        ("Avg total inflows" if is_average_view else "Total inflows", zar(income_total), "Salary + other personal income", "good"),
        ("Avg expenses" if is_average_view else "Expenses", zar(expense_total), f"{pct(expense_ratio)} of total inflows", "bad"),
        ("Avg saved / transferred" if is_average_view else "Saved / invested", zar(savings_total), f"Savings rate {pct(savings_rate)}", "blue"),
        ("Avg monthly deficit" if is_average_view else "Final monthly surplus", zar(net), "After expenses and savings", "good" if net >= 0 else "bad"),
    ])

    st.markdown(
        """
        <div class="small-note">
        The salary card now uses <b>cheque credit / net salary</b> only. Side income and investment income are shown separately so the dashboard does not confuse gross salary, net salary, and total monthly inflows.
        </div>
        """,
        unsafe_allow_html=True,
    )

    left, right = st.columns([1.05, 1], gap="large")

    with left:
        st.subheader("Average checkpoint summary" if is_average_view else "Monthly checkpoint summary")
        summary_fig = go.Figure()
        summary_fig.add_bar(name="Net salary", x=["Net salary"], y=[salary_total], marker_color="#16a34a", text=[zar(salary_total)], textposition="outside")
        summary_fig.add_bar(name="Other income", x=["Other income"], y=[other_income_total], marker_color="#3b82f6", text=[zar(other_income_total)], textposition="outside")
        summary_fig.add_bar(name="Expenses", x=["Expenses"], y=[expense_total], marker_color="#dc2626", text=[zar(expense_total)], textposition="outside")
        summary_fig.add_bar(name="Saved", x=["Saved"], y=[savings_total], marker_color="#2563eb", text=[zar(savings_total)], textposition="outside")
        summary_fig.add_bar(name="Final surplus", x=["Surplus"], y=[net], marker_color="#d97706", text=[zar(net)], textposition="outside")
        summary_fig.update_layout(barmode="group", showlegend=False)
        st.plotly_chart(responsive_chart_layout(summary_fig, height=330, show_legend=False), use_container_width=True, config={"responsive": True})

    with right:
        st.subheader("Average expense breakdown" if is_average_view else "Expense breakdown")
        exp_df = df[df["entry_type"] == "expense"]
        if exp_df.empty:
            st.info("No expenses entered for this view.")
        else:
            exp_by_cat = exp_df.groupby("category", as_index=False)["amount"].sum().sort_values("amount", ascending=False)
            exp_by_cat["amount"] = exp_by_cat["amount"] / period_divisor
            fig = px.pie(exp_by_cat, values="amount", names="category", hole=0.42)
            fig.update_traces(textposition="inside", textinfo="percent", hovertemplate="%{label}<br>R %{value:,.0f}<br>%{percent}<extra></extra>")
            fig.update_layout(height=330, margin=dict(l=10, r=10, t=30, b=10), legend=dict(orientation="v", y=0.5, x=1.02))
            st.plotly_chart(fig, use_container_width=True, config={"responsive": True})

    st.divider()

    # ─── Compact monthly summary table ──────────────────────────────────────────
    st.subheader("📊 Monthly Summary (All Months)")
    st.caption("Income, expenses, savings, and net cash flow for each month. Click column headers to sort.")
    
    # Load all checkpoints for the compact table
    all_checkpoints = get_checkpoint_df()
    if not all_checkpoints.empty:
        compact_summary = all_checkpoints.rename(columns={
            "month": "Month", 
            "income": "Income (R)", 
            "expenses": "Expenses (R)",
            "savings": "Savings (R)", 
            "net": "Net (R)",
        })[["Month", "Income (R)", "Expenses (R)", "Savings (R)", "Net (R)"]]
        
        st.dataframe(
            compact_summary,
            use_container_width=True,
            hide_index=True,
            column_config=money_column_config("Income (R)", "Expenses (R)", "Savings (R)", "Net (R)"),
        )
    else:
        st.info("No checkpoint data available yet.")
    
    st.divider()

    tab1, tab2, tab3 = st.tabs(["💵 Income", "💸 Expenses", "🏦 Savings & Investments"])

    with tab1:
        inc_df = df[df["entry_type"] == "income"]
        if inc_df.empty:
            st.info("No income entries.")
        else:
            if is_average_view:
                show = inc_df.groupby(["category", "subcategory"], as_index=False)["amount"].sum()
                show["amount"] = show["amount"] / period_divisor
                show["notes"] = period_label
            else:
                show = inc_df[["category", "subcategory", "amount", "notes"]]
            show = show.rename(columns={"category": "Category", "subcategory": "Detail", "amount": "Average / Amount (R)" if is_average_view else "Amount (R)", "notes": "Notes"})
            st.dataframe(show, use_container_width=True, hide_index=True, column_config=money_column_config("Amount (R)", "Average / Amount (R)"))
            st.caption(("Average net income received" if is_average_view else "Total net income received") + f": **{zar(income_total)}**")

    with tab2:
        if exp_df.empty:
            st.info("No expense entries.")
        else:
            fixed_df = exp_df[exp_df["category"].isin(DEFAULT_FIXED)]
            variable_df = exp_df[~exp_df["category"].isin(DEFAULT_FIXED)]
            for title, frame in [("Fixed expenses", fixed_df), ("Variable expenses", variable_df)]:
                if not frame.empty:
                    st.markdown(f"**{title}**")
                    if is_average_view:
                        show = frame.groupby("category", as_index=False)["amount"].sum()
                        show["amount"] = show["amount"] / period_divisor
                        comp = frame.groupby("category")["subcategory"].agg(
                            lambda s: ", ".join(s.value_counts().head(2).index.astype(str))[:70]
                        )
                        nun = frame.groupby("category")["subcategory"].nunique()
                        cnt = frame.groupby("category")["subcategory"].size()
                        show["subcategory"] = [
                            f"{c} (+{u - 2} more)" if u > 2 else c
                            for c, u in zip(comp, nun)
                        ]
                        show["notes"] = [f"{n} transactions · {period_label}" for n in cnt]
                    else:
                        show = frame[["category", "subcategory", "amount", "notes"]]
                    show = show.rename(columns={"category": "Category", "subcategory": "Detail", "amount": "Average / Amount (R)" if is_average_view else "Amount (R)", "notes": "Notes"})
                    st.dataframe(show, use_container_width=True, hide_index=True, column_config=money_column_config("Amount (R)", "Average / Amount (R)"))
                    st.caption(f"{title} {'average' if is_average_view else 'total'}: **{zar(frame['amount'].sum() / period_divisor)}**")
            st.caption(("Average expenses" if is_average_view else "Total expenses") + f": **{zar(expense_total)}**")

    with tab3:
        sav_df = df[df["entry_type"] == "savings"]
        if sav_df.empty:
            st.info("No savings entries.")
        else:
            if is_average_view:
                show = sav_df.groupby("category", as_index=False)["amount"].sum()
                show["amount"] = show["amount"] / period_divisor
                top = sav_df["subcategory"].mode()
                show["subcategory"] = str(top.iloc[0])[:60] if not top.empty else ""
                show["notes"] = f"{len(sav_df)} transactions · {period_label}"
            else:
                show = sav_df[["category", "subcategory", "amount", "notes"]]
            show = show.rename(columns={"category": "Category", "subcategory": "Detail", "amount": "Average / Amount (R)" if is_average_view else "Amount (R)", "notes": "Notes"})
            st.dataframe(show, use_container_width=True, hide_index=True, column_config=money_column_config("Amount (R)", "Average / Amount (R)"))
            st.caption(("Average saved/transferred" if is_average_view else "Total saved/invested") + f": **{zar(savings_total)}**")

# ─── Add Entries ──────────────────────────────────────────────────────────────
elif page == "➕ Add Entries":
    st.markdown('<div class="hero"><div class="hero-title">➕ Add Budget Entries</div><div class="hero-subtitle">Enter net salary / cheque credit as income. Use expense rows for spending and savings rows for money deliberately set aside.</div></div>', unsafe_allow_html=True)

    c1, c2 = st.columns([1.2, 2])
    with c1:
        entry_month = st.text_input("Month (YYYY-MM)", value=date.today().strftime("%Y-%m"))
    with c2:
        entry_type = st.radio("Type", ["income", "expense", "savings"], horizontal=True)

    categories = DEFAULT_INCOME if entry_type == "income" else DEFAULT_FIXED + DEFAULT_VARIABLE if entry_type == "expense" else DEFAULT_SAVINGS

    with st.form("single_entry_form", clear_on_submit=True):
        f1, f2 = st.columns(2)
        with f1:
            category = st.selectbox("Category", categories)
            amount = st.number_input("Amount (R)", min_value=0.0, step=100.0, format="%.2f")
        with f2:
            subcategory = st.text_input("Detail / Subcategory", placeholder="e.g. Cheque credit, Checkers, Petrol")
            notes = st.text_input("Notes", placeholder="Optional")
        submitted = st.form_submit_button("✅ Add Entry", use_container_width=True, type="primary")
        if submitted:
            if amount <= 0:
                st.error("Amount must be greater than 0.")
            else:
                add_entry(entry_month, category, subcategory, amount, entry_type, notes, "manual", 0)
                st.success(f"Added {category}: {zar(amount)} for {entry_month}")

    st.divider()
    st.subheader("📥 Quick import")
    st.caption("CSV format: category, subcategory, amount, type, notes. Example: Cheque Credit / Net Salary, Salary paid in, 33000, income, bank cheque credit")
    csv_data = st.text_area("Paste entries", height=150)
    if st.button("Import pasted entries"):
        imported = 0
        for raw_line in [l.strip() for l in csv_data.splitlines() if l.strip()]:
            parts = [p.strip() for p in raw_line.split(",")]
            if len(parts) < 3:
                st.warning(f"Skipped incomplete line: {raw_line}")
                continue
            try:
                amount = float(parts[2].replace("R", "").replace(" ", ""))
            except ValueError:
                st.warning(f"Skipped invalid amount: {raw_line}")
                continue
            add_entry(entry_month, parts[0], parts[1] if len(parts) > 1 else "", amount, parts[3] if len(parts) > 3 else entry_type, parts[4] if len(parts) > 4 else "", "manual", 0)
            imported += 1
        st.success(f"Imported {imported} entries")

# ─── Trends & Checkpoints ─────────────────────────────────────────────────────
elif page == "📈 Trends & Checkpoints":
    st.markdown('<div class="hero"><div class="hero-title">📈 Trends & Checkpoints</div><div class="hero-subtitle">Every point below is a visible monthly checkpoint value. No hidden/ambiguous line charts.</div></div>', unsafe_allow_html=True)

    checkpoints = get_checkpoint_df()
    if checkpoints.empty:
        st.info("No checkpoint data yet. Add at least one month of entries.")
        st.stop()

    latest = checkpoints.iloc[-1]
    render_metric_grid([
        ("Latest checkpoint", latest["month"], "Month currently used for trend comparison", "blue"),
        ("Total inflows", zar(latest["income"]), "Salary + other income received", "good"),
        ("Expenses", zar(latest["expenses"]), f"Expense ratio {pct(latest['expense_ratio'])}", "bad"),
        ("Saved", zar(latest["savings"]), f"Savings rate {pct(latest['savings_rate'])}", "blue"),
        ("Final surplus", zar(latest["net"]), "After expenses and savings", "good" if latest["net"] >= 0 else "bad"),
    ])

    st.subheader("Checkpoint table")
    checkpoint_show = checkpoints.rename(columns={
        "month": "Month", "income": "Total Inflows (R)", "expenses": "Expenses (R)", "savings": "Saved (R)",
        "cash_available": "After Expenses (R)", "net": "Final Surplus (R)", "savings_rate": "Savings Rate %", "expense_ratio": "Expense Ratio %"
    })
    st.dataframe(
        checkpoint_show,
        use_container_width=True,
        hide_index=True,
        column_config=money_column_config(
            "Total Inflows (R)",
            "Expenses (R)",
            "Saved (R)",
            "After Expenses (R)",
            "Final Surplus (R)",
        ),
    )

    if len(checkpoints) < 2:
        st.markdown('<div class="warning-note">Need at least two monthly checkpoints before a trend line is meaningful. With one month, the table above is the only honest view.</div>', unsafe_allow_html=True)
        st.stop()

    st.subheader("Monthly summary table")
    summary_show = checkpoints.rename(columns={
        "month": "Month", "income": "Income (R)", "expenses": "Expenses (R)",
        "savings": "Savings (R)", "net": "Net (R)",
    })[["Month", "Income (R)", "Expenses (R)", "Savings (R)", "Net (R)"]]
    st.dataframe(
        summary_show,
        use_container_width=True,
        hide_index=True,
        column_config=money_column_config("Income (R)", "Expenses (R)", "Savings (R)", "Net (R)"),
    )

    st.subheader("Trend lines with visible checkpoint values")
    metric_choice = st.multiselect(
        "Show trend lines",
        ["income", "expenses", "savings", "net"],
        default=["income", "expenses", "savings", "net"],
        format_func=lambda x: {"income": "Total inflows", "expenses": "Expenses", "savings": "Saved", "net": "Final surplus"}[x],
    )
    colors = {"income": "#16a34a", "expenses": "#dc2626", "savings": "#2563eb", "net": "#d97706"}
    names = {"income": "Total inflows", "expenses": "Expenses", "savings": "Saved", "net": "Final surplus"}

    fig = go.Figure()
    for col in metric_choice:
        fig.add_trace(
            go.Scatter(
                x=checkpoints["month"],
                y=checkpoints[col],
                name=names[col],
                mode="lines+markers+text",
                text=value_text(checkpoints[col]),
                textposition="top center",
                line=dict(color=colors[col], width=3),
                marker=dict(size=9),
                hovertemplate=f"{names[col]}<br>%{{x}}<br>R %{{y:,.0f}}<extra></extra>",
            )
        )
    st.plotly_chart(responsive_chart_layout(fig, height=430), use_container_width=True, config={"responsive": True})

    st.subheader("Month-over-month movement")
    latest = checkpoints.iloc[-1]
    prev = checkpoints.iloc[-2]
    render_metric_grid([
        ("Income movement", zar(latest["income"] - prev["income"]), f"{latest['month']} vs {prev['month']}", "good" if latest["income"] >= prev["income"] else "bad"),
        ("Expense movement", zar(latest["expenses"] - prev["expenses"]), "Lower is better", "bad" if latest["expenses"] > prev["expenses"] else "good"),
        ("Savings movement", zar(latest["savings"] - prev["savings"]), "Higher is better", "good" if latest["savings"] >= prev["savings"] else "bad"),
        ("Surplus movement", zar(latest["net"] - prev["net"]), "Higher is better", "good" if latest["net"] >= prev["net"] else "bad"),
    ])

    if len(checkpoints) >= 3:
        st.subheader("Simple 3-month checkpoint forecast")
        forecast_rows = []
        last_month = checkpoints["month"].iloc[-1]
        income_avg = checkpoints["income"].tail(3).mean()
        expenses_avg = checkpoints["expenses"].tail(3).mean()
        savings_avg = checkpoints["savings"].tail(3).mean()
        for i in range(1, 4):
            fm = datetime.strptime(last_month, "%Y-%m") + relativedelta(months=i)
            forecast_rows.append({
                "Month": fm.strftime("%Y-%m"),
                "Est. Total Inflows (R)": income_avg,
                "Est. Expenses (R)": expenses_avg,
                "Est. Saved (R)": savings_avg,
                "Est. Final Surplus (R)": income_avg - expenses_avg - savings_avg,
            })
        st.dataframe(
            pd.DataFrame(forecast_rows),
            use_container_width=True,
            hide_index=True,
            column_config=money_column_config(
                "Est. Total Inflows (R)",
                "Est. Expenses (R)",
                "Est. Saved (R)",
                "Est. Final Surplus (R)",
            ),
        )
        st.caption("Forecast uses the last three visible checkpoints only. It improves as more real months are added.")
    else:
        st.markdown('<div class="warning-note">Forecast disabled until three monthly checkpoints exist. Two months can show direction, but not a reliable forecast.</div>', unsafe_allow_html=True)

# ─── Data Table ───────────────────────────────────────────────────────────────
elif page == "📋 Data Table":
    st.markdown('<div class="hero"><div class="hero-title">📋 All Budget Data</div><div class="hero-subtitle">Audit and clean entries. Amounts shown here are what drive the checkpoints.</div></div>', unsafe_allow_html=True)
    all_df = get_all_entries()
    if all_df.empty:
        st.info("No data yet.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            # Newest-first picker; get_months() stays ascending for other consumers.
            filter_month = st.multiselect("Month", get_months()[::-1])
        with c2:
            filter_type = st.multiselect("Type", ["income", "expense", "savings"])
        display_df = all_df.copy()
        if filter_month:
            display_df = display_df[display_df["month"].isin(filter_month)]
        if filter_type:
            display_df = display_df[display_df["entry_type"].isin(filter_type)]
        show = display_df[["id", "month", "entry_type", "category", "subcategory", "amount", "notes", "source", "is_sample"]].rename(columns={
            "id": "ID", "month": "Month", "entry_type": "Type", "category": "Category", "subcategory": "Detail", "amount": "Amount (R)", "notes": "Notes", "source": "Source", "is_sample": "Sample?"
        })
        st.dataframe(show, use_container_width=True, hide_index=True, column_config=money_column_config("Amount (R)", "Average / Amount (R)"))
        st.divider()
        del_id = st.number_input("Entry ID to delete", min_value=1, step=1)
        if st.button("Delete entry"):
            delete_entry(int(del_id))
            st.success(f"Deleted entry {del_id}. Refresh if the table does not update immediately.")

# ─── Settings ─────────────────────────────────────────────────────────────────
elif page == "⚙️ Settings":
    st.markdown('<div class="hero"><div class="hero-title">⚙️ Settings</div><div class="hero-subtitle">Dashboard definitions and database status.</div></div>', unsafe_allow_html=True)
    st.subheader("Definitions")
    st.markdown("- **Net income received** = actual salary deposits in the cheque-credit account, not the workbook Salary sheet and not business income/other credits.")
    st.markdown("- **Expenses** = money spent or committed during the month.")
    st.markdown("- **Savings & investments** = deliberate transfers out of available cash. Current workbook rule: the R1,400 scheduled monthly transfer is savings; PSG AML remains a spending/outflow category because it is paid every month.")
    st.markdown("- **Checkpoint** = month-end totals rebuilt from the entries table and stored in `monthly_snapshots`.")
    st.markdown("- **Source** = where the entry came from. Current visible data is marked `seeded_sample`, `user_file`, or `manual`.")
    st.markdown("- **Sample?** = whether the row is placeholder data I generated, not your real financial record.")

    st.subheader("Default categories")
    st.markdown("**Income:** " + ", ".join(DEFAULT_INCOME))
    st.markdown("**Fixed expenses:** " + ", ".join(DEFAULT_FIXED))
    st.markdown("**Variable expenses:** " + ", ".join(DEFAULT_VARIABLE))
    st.markdown("**Savings:** " + ", ".join(DEFAULT_SAVINGS))

    st.subheader("Database")
    st.caption(f"Database path: `{DB_PATH}`")
    if st.button("Rebuild checkpoint table"):
        rebuild_monthly_checkpoints()
        st.success("Monthly checkpoints rebuilt from current budget entries.")
    conn = get_conn()
    entry_count = pd.read_sql("SELECT COUNT(*) AS cnt FROM budget_entries", conn)["cnt"].iloc[0]
    month_count = pd.read_sql("SELECT COUNT(DISTINCT month) AS cnt FROM budget_entries", conn)["cnt"].iloc[0]
    sample_count = pd.read_sql("SELECT COUNT(*) AS cnt FROM budget_entries WHERE is_sample=1", conn)["cnt"].iloc[0]
    real_count = pd.read_sql("SELECT COUNT(*) AS cnt FROM budget_entries WHERE source='user_file'", conn)["cnt"].iloc[0]
    snapshot_count = pd.read_sql("SELECT COUNT(*) AS cnt FROM monthly_snapshots", conn)["cnt"].iloc[0]
    conn.close()
    st.write(f"Entries: **{entry_count}**")
    st.write(f"Months: **{month_count}**")
    st.write(f"Real rows from user file: **{real_count}**")
    st.write(f"Seeded sample entries: **{sample_count}**")
    st.write(f"Checkpoints: **{snapshot_count}**")
