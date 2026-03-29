import streamlit as st
import pandas as pd
import altair as alt
from io import BytesIO
import os

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FACULTY_MASTER_PATH = os.path.join(_BASE_DIR, "data", "faculty_master.xlsx")
AWARDS_DATA_PATH = os.path.join(_BASE_DIR, "data", "awards_df.xls")
PROPOSALS_DATA_PATH = os.path.join(_BASE_DIR, "data", "proposals_df.xls")

AWARD_DATE_COL = "Award Finalize Date"
AWARD_AMOUNT_COL = "Award Obligated Total Cost"
AWARD_SPONSOR_COL = "Award Sponsor Name"
AWARD_TRANS_TYPE_COL = "Award Transaction Type Description"

PROPOSAL_DATE_COL = "Proposal Process Date"
PROPOSAL_AMOUNT_COL = "Proposal Total Cost"
PROPOSAL_FUNDED_FLAG_COL = "Proposal Funded Flag"
PROP_SPONSOR_COL = "Proposal Sponsor Name"

FACULTY_ID_COL = "Award PI Campus ID"
FACULTY_DEPT_COL = "Department"

QUARTER_SORT = ["Q1", "Q2", "Q3", "Q4"]

@st.cache_data
def load_excel_or_csv(path):
    if not os.path.exists(path):
        return None
    try:
        return pd.read_excel(path)
    except Exception:
        try:
            return pd.read_excel(path, engine='xlrd')
        except Exception:
            try:
                return pd.read_csv(path, sep=None, engine='python', on_bad_lines='skip')
            except Exception:
                return None

def normalize_funded_flag(flag_col):
    return flag_col.astype(str).str.lower().str.strip().isin(['y', 'yes', 'true', '1', 'funded'])

def collapse_nih_sponsors(name):
    if pd.isna(name):
        return name

    name_upper = str(name).upper().strip()

    nih_keywords = [
        "NATIONAL INSTITUTE",
        "NATIONAL INST",
        "NIH",
        "NATIONAL CANCER",
        "NATIONAL EYE",
        "LIBRARY OF MEDICINE",
        "FOGARTY",
        "NIMH",
        "CLINICAL CENTER",
        "ALCOHOL ABUSE",
        "ARTHRITIS, MUSCULOSKELETAL",
        "CHILD HEALTH & HUMAN",
        "DEAFNESS & OTHER COMMUNICATION",
        "DENTAL AND CRANIOFACIAL",
        "DRUG ABUSE",
        "ENVIRONMENTAL HEALTH SCIENCES",
        "NATIONAL CENTER FOR COMPLEMENTARY",
        "NATIONAL HEART, LUNG AND BLOOD",
        "NATIONAL HUMAN GENOME",
        "DIABETES AND DIGESTIVE",
        "NEUROLOGICAL DISORDERS & STROKE",
        "NURSING RESEARCH",
        "OFFICE OF THE DIRECTOR",
        "CENTER FOR SCIENTIFIC REVIEW"
    ]

    if any(k in name_upper for k in nih_keywords):
        return "NIH"

    return name

def get_fiscal_quarter(df, date_col):
    df = df.copy()
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    df = df.dropna(subset=[date_col])
    df["Fiscal Year"] = df[date_col].dt.year + (df[date_col].dt.month >= 7).astype(int)
    df["Quarter"] = ((df[date_col].dt.month - 7) % 12 // 3) + 1
    df["Fiscal Quarter"] = "Q" + df["Quarter"].astype(str)
    return df

st.set_page_config(page_title="BioSci Research Dashboard", layout="wide")
st.title("Research Development Portfolio Tool")

faculty_master = load_excel_or_csv(FACULTY_MASTER_PATH)
raw_awards = load_excel_or_csv(AWARDS_DATA_PATH)
raw_proposals = load_excel_or_csv(PROPOSALS_DATA_PATH)

if faculty_master is None:
    st.error(f"Required data file '{FACULTY_MASTER_PATH}' not found in /data.")
    st.stop()

faculty_master[FACULTY_ID_COL] = pd.to_numeric(faculty_master[FACULTY_ID_COL], errors="coerce").fillna(0).astype(int)
depts_and_ids = faculty_master[[FACULTY_ID_COL, FACULTY_DEPT_COL]].drop_duplicates()
campus_ids = depts_and_ids[FACULTY_ID_COL].tolist()

def process_df(df, date_col, sponsor_col, id_col_name=FACULTY_ID_COL):
    if df is None: return None
    if id_col_name in df.columns:
        df[FACULTY_ID_COL] = pd.to_numeric(df[id_col_name], errors="coerce").fillna(0).astype(int)
    elif "Proposal PI Campus ID" in df.columns:
        df[FACULTY_ID_COL] = pd.to_numeric(df["Proposal PI Campus ID"], errors="coerce").fillna(0).astype(int)

    df = df[df[FACULTY_ID_COL].isin(campus_ids)].copy()
    df = df.merge(depts_and_ids, on=FACULTY_ID_COL, how="left")

    if sponsor_col in df.columns:
        df[sponsor_col] = df[sponsor_col].apply(collapse_nih_sponsors)

    return get_fiscal_quarter(df, date_col)

final_awards = process_df(raw_awards, AWARD_DATE_COL, AWARD_SPONSOR_COL)
final_proposals = process_df(raw_proposals, PROPOSAL_DATE_COL, PROP_SPONSOR_COL)

# --- SIDEBAR FILTERS ---

all_fys = sorted(set(
    (final_awards["Fiscal Year"].unique().tolist() if final_awards is not None else []) +
    (final_proposals["Fiscal Year"].unique().tolist() if final_proposals is not None else [])
))

st.sidebar.header("Global Filters")

# Fiscal Year Selection: Individual or Start Year
st.sidebar.subheader("Fiscal Year Selection")
fy_mode = st.sidebar.radio("Mode", ["Start year", "Individual selection"], label_visibility="collapsed")

if fy_mode == "Individual selection":
    selected_fys = st.sidebar.multiselect("Fiscal Year", options=sorted(all_fys, reverse=True))
else:
    default_start = max(all_fys) - 5 if all_fys else None
    start_year = st.sidebar.selectbox(
        "Show from FY",
        options=sorted(all_fys),
        index=all_fys.index(default_start) if default_start in all_fys else 0
    )
    selected_fys = [fy for fy in all_fys if fy >= start_year]

dept_values = sorted(depts_and_ids[FACULTY_DEPT_COL].unique().tolist())
selected_depts = st.sidebar.multiselect("Department", options=dept_values)

faculty_names_map = faculty_master.set_index(FACULTY_ID_COL)["Name"].to_dict()
faculty_options = sorted([(fid, faculty_names_map.get(fid, f"ID: {fid}")) for fid in campus_ids], key=lambda x: x[1])
selected_faculty_info = st.sidebar.multiselect("Faculty PI", options=faculty_options, format_func=lambda x: x[1])
selected_faculty_ids = [x[0] for x in selected_faculty_info]

with st.sidebar.expander("Admin: Refresh Data"):
    st.caption("Upload new files to override hard-coded data.")
    up_awd = st.file_uploader("Awards", type=["xls", "xlsx", "csv"])
    up_prop = st.file_uploader("Proposals", type=["xls", "xlsx", "csv"])

def apply_filters(df, is_award=False):
    if df is None: return None
    out = df.copy()
    if is_award and AWARD_TRANS_TYPE_COL in out.columns:
        out = out[out[AWARD_TRANS_TYPE_COL].isin(["New", "Renewal", "Supplement"])]
    if selected_fys:
        out = out[out["Fiscal Year"].isin(selected_fys)]
    if selected_depts:
        out = out[out[FACULTY_DEPT_COL].isin(selected_depts)]
    if selected_faculty_ids:
        out = out[out[FACULTY_ID_COL].isin(selected_faculty_ids)]
    return out

fa = apply_filters(final_awards, is_award=True)
fp = apply_filters(final_proposals)

tab_overview, tab_faculty, tab_tables = st.tabs(["Portfolio Overview", "Faculty Drill-Down", "Raw Data"])

with tab_overview:
    view_mode = st.radio("Metric Mode", ["Award View", "Proposal View"], horizontal=True)

    # Visualization toggle (only meaningful when multiple departments are selected)
    multi_dept = len(selected_depts) > 1
    if multi_dept:
        viz_mode = st.radio(
            "Visualization",
            ["Aggregated", "Side-by-Side by Department"],
            horizontal=True,
            key="viz_mode"
        )
    else:
        viz_mode = "Aggregated"

    quarter_color = alt.Color(
        "Fiscal Quarter:O",
        sort=QUARTER_SORT,
        scale=alt.Scale(scheme="blues", reverse=True)
    )
    quarter_order = alt.Order("Quarter:O", sort="ascending")

    if view_mode == "Award View" and fa is not None:
        total_awd = pd.to_numeric(fa[AWARD_AMOUNT_COL], errors="coerce").sum()
        c1, c2 = st.columns(2)
        c1.metric("Award Count", f"{len(fa):,}")
        c2.metric("Total Obligated", f"${total_awd:,.0f}")

        if viz_mode == "Side-by-Side by Department":
            awd_time = fa.groupby([FACULTY_DEPT_COL, "Fiscal Year", "Fiscal Quarter"], as_index=False).agg(
                Count=("Fiscal Year", "size"), Dollars=(AWARD_AMOUNT_COL, "sum")
            )
            col1, col2 = st.columns(2)
            with col1:
                chart = alt.Chart(awd_time).mark_bar().encode(
                    x=alt.X("Fiscal Year:O", title="FY"),
                    y=alt.Y("Count:Q"),
                    color=quarter_color,
                    order=quarter_order,
                    column=alt.Column(f"{FACULTY_DEPT_COL}:N", header=alt.Header(labelOrient="bottom", titleOrient="bottom")),
                    tooltip=[FACULTY_DEPT_COL, "Fiscal Year", "Fiscal Quarter", "Count"]
                ).properties(title="Award Count", width=150, height=300)
                st.altair_chart(chart)
            with col2:
                chart = alt.Chart(awd_time).mark_bar().encode(
                    x=alt.X("Fiscal Year:O", title="FY"),
                    y=alt.Y("Dollars:Q"),
                    color=quarter_color,
                    order=quarter_order,
                    column=alt.Column(f"{FACULTY_DEPT_COL}:N", header=alt.Header(labelOrient="bottom", titleOrient="bottom")),
                    tooltip=[FACULTY_DEPT_COL, "Fiscal Year", "Fiscal Quarter", alt.Tooltip("Dollars:Q", format="$,.0f")]
                ).properties(title="Award Dollars", width=150, height=300)
                st.altair_chart(chart)
        else:
            awd_time = fa.groupby(["Fiscal Year", "Fiscal Quarter"], as_index=False).agg(
                Count=("Fiscal Year", "size"), Dollars=(AWARD_AMOUNT_COL, "sum")
            )
            col1, col2 = st.columns(2)
            with col1:
                st.altair_chart(alt.Chart(awd_time).mark_bar().encode(
                    x="Fiscal Year:O", y="Count:Q", color=quarter_color, order=quarter_order,
                    tooltip=["Fiscal Year", "Fiscal Quarter", "Count"]
                ).properties(title="Award Count", height=350), use_container_width=True)
            with col2:
                st.altair_chart(alt.Chart(awd_time).mark_bar().encode(
                    x="Fiscal Year:O", y="Dollars:Q", color=quarter_color, order=quarter_order,
                    tooltip=["Fiscal Year", "Fiscal Quarter", alt.Tooltip("Dollars:Q", format="$,.0f")]
                ).properties(title="Award Dollars", height=350), use_container_width=True)

    elif view_mode == "Proposal View" and fp is not None:
        total_prop_val = pd.to_numeric(fp[PROPOSAL_AMOUNT_COL], errors="coerce").sum()
        funded_mask = normalize_funded_flag(fp[PROPOSAL_FUNDED_FLAG_COL]) if PROPOSAL_FUNDED_FLAG_COL in fp.columns else pd.Series(False, index=fp.index)
        n_submitted = len(fp)
        n_funded = int(funded_mask.sum())
        success_rate = (n_funded / n_submitted * 100) if n_submitted > 0 else 0.0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Proposals Submitted", f"{n_submitted:,}")
        c2.metric("Total Requested", f"${total_prop_val:,.0f}")
        c3.metric("Funded", f"{n_funded:,}")
        c4.metric("Success Rate", f"{success_rate:.1f}%")

        if viz_mode == "Side-by-Side by Department":
            prop_time = fp.groupby([FACULTY_DEPT_COL, "Fiscal Year", "Fiscal Quarter"], as_index=False).agg(
                Count=("Fiscal Year", "size"), Dollars=(PROPOSAL_AMOUNT_COL, "sum")
            )
            pcol1, pcol2 = st.columns(2)
            with pcol1:
                chart = alt.Chart(prop_time).mark_bar().encode(
                    x=alt.X("Fiscal Year:O", title="FY"),
                    y=alt.Y("Count:Q"),
                    color=quarter_color,
                    order=quarter_order,
                    column=alt.Column(f"{FACULTY_DEPT_COL}:N", header=alt.Header(labelOrient="bottom", titleOrient="bottom")),
                    tooltip=[FACULTY_DEPT_COL, "Fiscal Year", "Fiscal Quarter", "Count"]
                ).properties(title="Proposal Count", width=150, height=300)
                st.altair_chart(chart)
            with pcol2:
                chart = alt.Chart(prop_time).mark_bar().encode(
                    x=alt.X("Fiscal Year:O", title="FY"),
                    y=alt.Y("Dollars:Q"),
                    color=quarter_color,
                    order=quarter_order,
                    column=alt.Column(f"{FACULTY_DEPT_COL}:N", header=alt.Header(labelOrient="bottom", titleOrient="bottom")),
                    tooltip=[FACULTY_DEPT_COL, "Fiscal Year", "Fiscal Quarter", alt.Tooltip("Dollars:Q", format="$,.0f")]
                ).properties(title="Proposal Requested $", width=150, height=300)
                st.altair_chart(chart)
        else:
            prop_time = fp.groupby(["Fiscal Year", "Fiscal Quarter"], as_index=False).agg(
                Count=("Fiscal Year", "size"), Dollars=(PROPOSAL_AMOUNT_COL, "sum")
            )
            pcol1, pcol2 = st.columns(2)
            with pcol1:
                st.altair_chart(alt.Chart(prop_time).mark_bar().encode(
                    x="Fiscal Year:O", y="Count:Q", color=quarter_color, order=quarter_order,
                    tooltip=["Fiscal Year", "Fiscal Quarter", "Count"]
                ).properties(title="Proposal Count", height=350), use_container_width=True)
            with pcol2:
                st.altair_chart(alt.Chart(prop_time).mark_bar().encode(
                    x="Fiscal Year:O", y="Dollars:Q", color=quarter_color, order=quarter_order,
                    tooltip=["Fiscal Year", "Fiscal Quarter", alt.Tooltip("Dollars:Q", format="$,.0f")]
                ).properties(title="Proposal Requested $", height=350), use_container_width=True)

    st.divider()
    st.markdown("### Breakdown Analysis")
    b1, b2 = st.columns(2)
    df_src = fa if view_mode == "Award View" else fp
    spon_col = AWARD_SPONSOR_COL if view_mode == "Award View" else PROP_SPONSOR_COL
    is_proposal_view = view_mode == "Proposal View"

    with b1:
        st.write("**By Department**")
        if df_src is not None:
            if is_proposal_view and PROPOSAL_FUNDED_FLAG_COL in df_src.columns:
                dept_funded = normalize_funded_flag(df_src[PROPOSAL_FUNDED_FLAG_COL])
                dept_stats = df_src.groupby(FACULTY_DEPT_COL).agg(
                    Submitted=(FACULTY_DEPT_COL, "size")
                ).reset_index()
                dept_stats["Funded"] = df_src[dept_funded].groupby(df_src[FACULTY_DEPT_COL]).size().reindex(dept_stats[FACULTY_DEPT_COL]).values
                dept_stats["Funded"] = dept_stats["Funded"].fillna(0).astype(int)
                dept_stats["Success Rate"] = (dept_stats["Funded"] / dept_stats["Submitted"] * 100).round(1).astype(str) + "%"
                st.dataframe(dept_stats, hide_index=True, use_container_width=True)
            else:
                dept_stats = df_src.groupby(FACULTY_DEPT_COL).size().reset_index(name="Count")
                st.dataframe(dept_stats, hide_index=True, use_container_width=True)

    with b2:
        st.write("**By Sponsor (NIH Grouped)**")
        if df_src is not None:
            if is_proposal_view and PROPOSAL_FUNDED_FLAG_COL in df_src.columns:
                prop_funded = normalize_funded_flag(df_src[PROPOSAL_FUNDED_FLAG_COL])
                spon_stats = df_src.groupby(spon_col).agg(
                    Submitted=(spon_col, "size")
                ).reset_index()
                spon_stats["Funded"] = df_src[prop_funded].groupby(df_src[spon_col]).size().reindex(spon_stats[spon_col]).values
                spon_stats["Funded"] = spon_stats["Funded"].fillna(0).astype(int)
                spon_stats["Success Rate"] = (spon_stats["Funded"] / spon_stats["Submitted"] * 100).round(1).astype(str) + "%"
                spon_stats = spon_stats.sort_values("Submitted", ascending=False).head(10)
                st.dataframe(spon_stats, hide_index=True, use_container_width=True)
            else:
                spon_stats = df_src.groupby(spon_col).size().sort_values(ascending=False).head(10).reset_index(name="Count")
                st.dataframe(spon_stats, hide_index=True, use_container_width=True)

with tab_faculty:
    st.subheader("Faculty Activity Drill-Down")
    selected_faculty_drill = st.selectbox("Select Faculty Member", options=faculty_options, format_func=lambda x: x[1])
    fid = selected_faculty_drill[0]

    st.markdown(f"#### Master Activity Table: {selected_faculty_drill[1]} (ID: {int(fid)})")

    f_awds = fa[fa[FACULTY_ID_COL] == fid].copy() if fa is not None else pd.DataFrame()
    f_props = fp[fp[FACULTY_ID_COL] == fid].copy() if fp is not None else pd.DataFrame()

    # Proposal success rate for this faculty member
    if not f_props.empty and PROPOSAL_FUNDED_FLAG_COL in f_props.columns:
        fac_funded_mask = normalize_funded_flag(f_props[PROPOSAL_FUNDED_FLAG_COL])
        fac_submitted = len(f_props)
        fac_funded = int(fac_funded_mask.sum())
        fac_rate = (fac_funded / fac_submitted * 100) if fac_submitted > 0 else 0.0
        m1, m2, m3 = st.columns(3)
        m1.metric("Proposals Submitted", f"{fac_submitted:,}")
        m2.metric("Funded", f"{fac_funded:,}")
        m3.metric("Success Rate", f"{fac_rate:.1f}%")

    f_awds["Category"] = "AWARD"
    f_awds = f_awds.rename(columns={AWARD_DATE_COL: "Date", AWARD_SPONSOR_COL: "Sponsor", "Award Project Title": "Title", AWARD_AMOUNT_COL: "Amount"})

    f_props["Category"] = "PROPOSAL"
    if PROPOSAL_FUNDED_FLAG_COL in f_props.columns:
        f_props["Funded"] = normalize_funded_flag(f_props[PROPOSAL_FUNDED_FLAG_COL]).map({True: "Yes", False: "No"})
    f_props = f_props.rename(columns={PROPOSAL_DATE_COL: "Date", PROP_SPONSOR_COL: "Sponsor", "Proposal Project Title": "Title", PROPOSAL_AMOUNT_COL: "Amount"})

    master_activity = pd.concat([f_awds, f_props], axis=0, ignore_index=True)

    if not master_activity.empty:
        master_activity = master_activity.sort_values("Date", ascending=False)
        master_activity["Date"] = pd.to_datetime(master_activity["Date"]).dt.strftime('%Y-%m-%d')
        master_activity["Amount"] = master_activity["Amount"].apply(lambda x: f"${x:,.0f}" if pd.notnull(x) and x != 0 else "—")

        cols = ["Date", "Category", "Title", "Sponsor", "Amount", "Fiscal Year", "Fiscal Quarter"]
        if "Funded" in master_activity.columns:
            cols.append("Funded")
        st.dataframe(master_activity[cols], hide_index=True, use_container_width=True)
    else:
        st.info("No activity found for selected filters.")

with tab_tables:
    st.markdown("### Full Filtered Datasets")
    if fa is not None: st.dataframe(fa, hide_index=True)
    if fp is not None: st.dataframe(fp, hide_index=True)
