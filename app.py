"""
GENESIS & TIDE 2.0 — Startup Incubation Dashboard (Madhya Pradesh)

Run with:
    streamlit run app.py

Excel file required in the same folder:
    GENESIS & TIDE.xlsx
"""

import re
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


# ============================================================================
# PAGE CONFIGURATION
# ============================================================================

st.set_page_config(
    page_title="GENESIS & TIDE 2.0 Dashboard",
    page_icon="📊",
    layout="wide",
)


# ============================================================================
# CUSTOM CSS
# ============================================================================

st.markdown(
    """
    <style>

    /* Main page */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1450px;
    }

    /* Title */
    h1 {
        color: #1E2620;
        font-size: 2rem !important;
        font-weight: 700 !important;
        margin-bottom: 0.2rem !important;
    }

    h2, h3 {
        color: #1E2620;
        font-weight: 650 !important;
    }

    /* Caption */
    .stCaption {
        color: #666C65;
    }

    /* KPI cards */
    div[data-testid="stMetric"] {
        background: #FFFFFF;
        border: 1px solid #E4DED2;
        border-radius: 8px;
        padding: 16px 18px 14px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }

    [data-testid="stMetricLabel"],
    [data-testid="stMetricLabel"] * {
        font-size: 0.76rem;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        color: #5B6158 !important;
        -webkit-text-fill-color: #5B6158 !important;
    }

    [data-testid="stMetricValue"],
    [data-testid="stMetricValue"] * {
        font-size: 1.9rem;
        font-weight: 700;
        color: #1E2620 !important;
        -webkit-text-fill-color: #1E2620 !important;
    }

    /* Dataframe */
    div[data-testid="stDataFrame"] {
        border: 1px solid #E4DED2;
        border-radius: 6px;
    }

    /* Divider */
    hr {
        margin-top: 1.5rem;
        margin-bottom: 1.5rem;
        border-color: #E7E2D8;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================================
# CONSTANTS
# ============================================================================

GENESIS_COLOR = "#0B6E4F"
TIDE_COLOR = "#D9860F"

DEFAULT_FILE = Path(__file__).parent / "GENESIS & TIDE.xlsx"


# ============================================================================
# INCUBATOR NAME CANONICALIZATION
# ============================================================================

CANON_MAP = {

    "anupam innovation and incubation center a unit of kothiya society for sustainable development and research":
        "Anupam Innovation and Incubation Centre, Bhopal",

    "anupam innovation and incubation center":
        "Anupam Innovation and Incubation Centre, Bhopal",

    "anupam innovation and incubation centre, a unit of kothiya society for sutainable development and research, bhopal":
        "Anupam Innovation and Incubation Centre, Bhopal",

    "anupam innovation and incubation centre, a unit of kothiya society for sustainable development and research, bhopal":
        "Anupam Innovation and Incubation Centre, Bhopal",

    "iim udaipur incubation centre":
        "IIM Udaipur Incubation Centre",

    "iiti drishti cps foundation":
        "IITI Drishti CPS Foundation",

    "iimn foundation for entrepreneurship development":
        "IIMN Foundation for Entrepreneurship Development",

    "aic-rntu foundation":
        "AIC-RNTU Foundation",

    "ciie iit(ism) dhanbad foundation":
        "CIIE IIT(ISM) Dhanbad Foundation",

    "iit mandi catalyst":
        "IIT Mandi Catalyst",

    "technology innovation and incubation centre, iiitm-gwalior":
        "Technology Innovation and Incubation Centre, IIITM-Gwalior",
}


def canon(name: str) -> str:
    """
    Standardize incubator names to avoid duplicate names caused by
    spelling/capitalization variations.
    """

    key = re.sub(
        r"\s+",
        " ",
        str(name).strip().lower()
    ).rstrip(".")

    return CANON_MAP.get(
        key,
        re.sub(r"\s+", " ", str(name).strip())
    )


# ============================================================================
# DATA LOADING
# ============================================================================

@st.cache_data
def load_data(file_path):

    xls = pd.ExcelFile(file_path)

    # ------------------------------------------------------------------------
    # Read startup sheets
    # ------------------------------------------------------------------------

    def sheet_rows(
        sheet_name,
        incubator_col=1,
        startup_col=2,
        state_col=3
    ):

        df = pd.read_excel(
            xls,
            sheet_name=sheet_name,
            header=0
        )

        df = df.dropna(how="all")

        rows = []

        for _, r in df.iterrows():

            # Skip rows where incubator is missing
            inc = r.iloc[incubator_col]

            if pd.isna(inc):
                continue

            startup = (
                str(r.iloc[startup_col]).strip()
                if not pd.isna(r.iloc[startup_col])
                else ""
            )

            state = (
                str(r.iloc[state_col]).strip()
                if not pd.isna(r.iloc[state_col])
                else ""
            )

            rows.append(
                {
                    "incubator": canon(inc),
                    "startup": startup,
                    "state": state,
                }
            )

        return rows

    # ------------------------------------------------------------------------
    # GENESIS sheets
    # ------------------------------------------------------------------------

    genesis_sheets = {
        "EIR 1": "GENESIS EIR 1",
        "EIR 2": "GENESIS EIR 2",
        "Investment": "GENESIS INVESTMENT",
    }

    startup_records = []

    for eir_label, sheet_name in genesis_sheets.items():

        for row in sheet_rows(sheet_name):

            row.update(
                scheme="GENESIS",
                eir=eir_label
            )

            startup_records.append(row)

    # ------------------------------------------------------------------------
    # TIDE 2.0
    # ------------------------------------------------------------------------

    for row in sheet_rows("TIDE 2.0"):

        row.update(
            scheme="TIDE 2.0",
            eir=None
        )

        startup_records.append(row)

    startups = pd.DataFrame(startup_records)

    # ------------------------------------------------------------------------
    # Madhya Pradesh incubator sheet
    # ------------------------------------------------------------------------

    extra_sheet_name = next(
        (
            s
            for s in xls.sheet_names
            if "incubator in" in s.lower()
        ),
        None
    )

    inc_rows = []

    if extra_sheet_name:

        extra_df = pd.read_excel(
            xls,
            sheet_name=extra_sheet_name,
            header=0
        ).dropna(how="all")

        for _, r in extra_df.iterrows():

            name = r.iloc[0]
            scheme = r.iloc[1]

            if pd.isna(name) or pd.isna(scheme):
                continue

            inc_rows.append(
                {
                    "name": canon(name),
                    "scheme": str(scheme).strip()
                }
            )

    incubators = pd.DataFrame(inc_rows)

    if not incubators.empty:

        incubators = (
            incubators
            .drop_duplicates()
            .sort_values("name")
            .reset_index(drop=True)
        )

    return {
        "startups": startups,
        "incubators": incubators
    }


# ============================================================================
# CHECK EXCEL FILE
# ============================================================================

if not DEFAULT_FILE.exists():

    st.error(
        f"Couldn't find **{DEFAULT_FILE.name}**.\n\n"
        f"Please place the Excel file in the same folder as `app.py`."
    )

    st.stop()


# ============================================================================
# LOAD DATA
# ============================================================================

try:

    data = load_data(DEFAULT_FILE)

    rows = data["startups"]
    inc_rows = data["incubators"]

except Exception as e:

    st.error(
        "There was an error while reading the Excel file."
    )

    st.exception(e)

    st.stop()


# ============================================================================
# SAFETY CHECKS
# ============================================================================

if rows.empty:

    st.warning("No startup records were found in the Excel file.")

    st.stop()


# ============================================================================
# HEADER
# ============================================================================

st.title(
    "GENESIS & TIDE 2.0 — Startup Incubation Dashboard"
)

st.caption(
    "Startup and incubator activity under the GENESIS and TIDE 2.0 schemes, Madhya Pradesh."
)


# ============================================================================
# KPI CALCULATIONS
# ============================================================================

total_startups = len(rows)

genesis_startups = int(
    (rows["scheme"] == "GENESIS").sum()
)

tide_startups = int(
    (rows["scheme"] == "TIDE 2.0").sum()
)

mp_incubators = (
    inc_rows["name"].nunique()
    if not inc_rows.empty
    else 0
)


# ============================================================================
# KPI CARDS
# ============================================================================

c1, c2, c3, c4 = st.columns(4)

with c1:
    st.metric(
        "Total Startups",
        f"{total_startups:,}"
    )

with c2:
    st.metric(
        "GENESIS Startups",
        f"{genesis_startups:,}"
    )

with c3:
    st.metric(
        "TIDE 2.0 Startups",
        f"{tide_startups:,}"
    )

with c4:
    st.metric(
        "Madhya Pradesh Incubators",
        f"{mp_incubators:,}"
    )


st.divider()


# ============================================================================
# GENESIS EIR-WISE + SCHEME-WISE
# ============================================================================

col1, col2 = st.columns(2)


# ----------------------------------------------------------------------------
# GENESIS EIR-WISE
# ----------------------------------------------------------------------------

with col1:

    st.subheader("GENESIS — EIR-wise startup count")

    genesis_rows = rows[
        rows["scheme"] == "GENESIS"
    ]

    eir_order = [
        "EIR 1",
        "EIR 2",
        "Investment"
    ]

    eir_counts = (
        genesis_rows["eir"]
        .value_counts()
        .reindex(eir_order, fill_value=0)
        .reset_index()
    )

    eir_counts.columns = [
        "EIR",
        "Startups"
    ]

    fig = px.bar(
        eir_counts,
        x="EIR",
        y="Startups",
        text="Startups",
        color_discrete_sequence=[GENESIS_COLOR],
    )

    fig.update_traces(
        textposition="outside",
        cliponaxis=False
    )

    fig.update_layout(
        showlegend=False,
        yaxis_title=None,
        xaxis_title=None,
        margin=dict(
            l=10,
            r=10,
            t=20,
            b=10
        ),
        plot_bgcolor="white",
        paper_bgcolor="white",
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )


# ----------------------------------------------------------------------------
# SCHEME-WISE
# ----------------------------------------------------------------------------

with col2:

    st.subheader("Scheme-wise startup comparison")

    scheme_counts = (
        rows["scheme"]
        .value_counts()
        .reindex(
            ["GENESIS", "TIDE 2.0"],
            fill_value=0
        )
        .reset_index()
    )

    scheme_counts.columns = [
        "Scheme",
        "Startups"
    ]

    fig = px.bar(
        scheme_counts,
        x="Scheme",
        y="Startups",
        color="Scheme",
        text="Startups",
        color_discrete_map={
            "GENESIS": GENESIS_COLOR,
            "TIDE 2.0": TIDE_COLOR,
        },
    )

    fig.update_traces(
        textposition="outside",
        cliponaxis=False
    )

    fig.update_layout(
        showlegend=False,
        yaxis_title=None,
        xaxis_title=None,
        margin=dict(
            l=10,
            r=10,
            t=20,
            b=10
        ),
        plot_bgcolor="white",
        paper_bgcolor="white",
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )


# ============================================================================
# INCUBATOR-WISE STARTUP DISTRIBUTION
# ============================================================================

st.divider()

st.subheader("Incubator-wise startup distribution")

incubator_counts = (
    rows
    .groupby(["incubator", "scheme"])
    .size()
    .reset_index(name="Startups")
)

incubator_total = (
    incubator_counts
    .groupby("incubator")["Startups"]
    .sum()
    .reset_index()
)

incubator_total = (
    incubator_total
    .sort_values(
        "Startups",
        ascending=False
    )
    .head(10)
)

fig = px.bar(
    incubator_total.sort_values(
        "Startups",
        ascending=True
    ),
    x="Startups",
    y="incubator",
    orientation="h",
    text="Startups",
    color_discrete_sequence=[GENESIS_COLOR],
)

fig.update_traces(
    textposition="outside",
    cliponaxis=False
)

fig.update_layout(
    xaxis_title=None,
    yaxis_title=None,
    showlegend=False,
    margin=dict(
        l=10,
        r=60,
        t=20,
        b=10
    ),
    plot_bgcolor="white",
    paper_bgcolor="white",
)

st.plotly_chart(
    fig,
    use_container_width=True
)

st.caption(
    "Top 10 incubators by total number of startups across GENESIS and TIDE 2.0."
)


# ============================================================================
# INCUBATOR-WISE SCHEME SUMMARY
# ============================================================================

st.divider()

st.subheader("Incubator-wise scheme summary")


# Create GENESIS/TIDE counts per incubator
scheme_summary = (
    rows
    .pivot_table(
        index="incubator",
        columns="scheme",
        values="startup",
        aggfunc="count",
        fill_value=0
    )
    .reset_index()
)

# Make sure both columns exist
if "GENESIS" not in scheme_summary.columns:
    scheme_summary["GENESIS"] = 0

if "TIDE 2.0" not in scheme_summary.columns:
    scheme_summary["TIDE 2.0"] = 0


scheme_summary["Total"] = (
    scheme_summary["GENESIS"]
    + scheme_summary["TIDE 2.0"]
)


# Sort by total startups
scheme_summary = scheme_summary.sort_values(
    "Total",
    ascending=False
)


# Rename columns
scheme_summary = scheme_summary.rename(
    columns={
        "incubator": "Incubator",
        "GENESIS": "GENESIS",
        "TIDE 2.0": "TIDE 2.0",
        "Total": "Total Startups",
    }
)


# Convert to integer
for col in [
    "GENESIS",
    "TIDE 2.0",
    "Total Startups"
]:

    scheme_summary[col] = (
        scheme_summary[col]
        .fillna(0)
        .astype(int)
    )


st.dataframe(
    scheme_summary[
        [
            "Incubator",
            "GENESIS",
            "TIDE 2.0",
            "Total Startups"
        ]
    ],
    use_container_width=True,
    hide_index=True,
    column_config={
        "Incubator": st.column_config.TextColumn(
            "Incubator"
        ),
        "GENESIS": st.column_config.NumberColumn(
            "GENESIS",
            format="%d"
        ),
        "TIDE 2.0": st.column_config.NumberColumn(
            "TIDE 2.0",
            format="%d"
        ),
        "Total Startups": st.column_config.NumberColumn(
            "Total Startups",
            format="%d"
        ),
    },
)


# ============================================================================
# EIR + SCHEME SUMMARY
# ============================================================================

st.divider()

st.subheader("GENESIS & TIDE 2.0 — Summary")


# GENESIS EIR counts
eir_summary = (
    rows[rows["scheme"] == "GENESIS"]
    .groupby("eir")
    .size()
    .reindex(
        [
            "EIR 1",
            "EIR 2",
            "Investment"
        ],
        fill_value=0
    )
)


summary_table = pd.DataFrame(
    {
        "Category": [
            "GENESIS — EIR 1",
            "GENESIS — EIR 2",
            "GENESIS — Investment",
            "GENESIS — Total",
            "TIDE 2.0 — Total",
            "GENESIS + TIDE 2.0 — Total",
        ],
        "Startups": [
            int(eir_summary["EIR 1"]),
            int(eir_summary["EIR 2"]),
            int(eir_summary["Investment"]),
            genesis_startups,
            tide_startups,
            total_startups,
        ],
    }
)


st.dataframe(
    summary_table,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Category": st.column_config.TextColumn(
            "Category"
        ),
        "Startups": st.column_config.NumberColumn(
            "Startups",
            format="%d"
        ),
    },
)


# ============================================================================
# MADHYA PRADESH INCUBATORS
# ============================================================================

st.divider()

st.subheader("Madhya Pradesh incubators")


if inc_rows.empty:

    st.caption(
        "No Madhya Pradesh incubator records were found."
    )

else:

    # ------------------------------------------------------------------------
    # Calculate startup counts for each incubator
    # ------------------------------------------------------------------------

    startup_count_by_incubator = (
        rows
        .groupby("incubator")
        .size()
        .reset_index(name="Startup Count")
    )

    # ------------------------------------------------------------------------
    # Merge startup counts with MP incubator list
    # ------------------------------------------------------------------------

    incubator_detail = inc_rows.copy()

    incubator_detail = incubator_detail.merge(
        startup_count_by_incubator,
        left_on="name",
        right_on="incubator",
        how="left"
    )

    incubator_detail["Startup Count"] = (
        incubator_detail["Startup Count"]
        .fillna(0)
        .astype(int)
    )

    # ------------------------------------------------------------------------
    # Remove duplicate incubator records
    # ------------------------------------------------------------------------

    incubator_detail = (
        incubator_detail[
            [
                "name",
                "scheme",
                "Startup Count"
            ]
        ]
        .drop_duplicates()
        .sort_values(
            ["Startup Count", "name"],
            ascending=[False, True]
        )
    )

    incubator_detail = incubator_detail.rename(
        columns={
            "name": "Incubator",
            "scheme": "Scheme",
        }
    )

    # ------------------------------------------------------------------------
    # Display
    # ------------------------------------------------------------------------

    st.dataframe(
        incubator_detail,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Incubator": st.column_config.TextColumn(
                "Incubator"
            ),
            "Scheme": st.column_config.TextColumn(
                "Scheme"
            ),
            "Startup Count": st.column_config.NumberColumn(
                "Startup Count",
                format="%d"
            ),
        },
    )

    st.caption(
        f"{inc_rows['name'].nunique()} unique Madhya Pradesh incubator(s) shown."
    )


# ============================================================================
# FOOTER
# ============================================================================

st.divider()

st.caption(
    "Source: GENESIS & TIDE 2.0 startup and incubator data provided in the source Excel file."
)
