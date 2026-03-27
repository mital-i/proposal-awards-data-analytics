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

def normalize_funded_flag(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip().str.lower()
    return s.isin(["y", "yes", "true", "1", "funded"])

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    return output.getvalue()

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

# --- 4. SIDEBAR FILTERS ---

all_fys = sorted(set(
    (final_awards["Fiscal Year"].unique().tolist() if final_awards is not None else []) +
    (final_proposals["Fiscal Year"].unique().tolist() if final_proposals is not None else [])
), reverse=True)

st.sidebar.header("Global Filters")

# Feature 2: Start-year selector
st.sidebar.markdown("**Fiscal Year Selection**")
year_selection_mode = st.sidebar.radio("Select years by:", options=["Individual selection", "Start year"], key="year_mode")

if year_selection_mode == "Start year":
    if all_fys:
        default_start = all_fys[-6] if len(all_fys) >= 6 else all_fys[-1]
        start_year = st.sidebar.selectbox(
            "Start Fiscal Year:",
            options=all_fys[::-1],
            index=list(all_fys[::-1]).index(default_start),
            key="start_year_select"
        )
        selected_fys = [fy for fy in all_fys if fy >= start_year]
    else:
        selected_fys = []
        st.sidebar.info("No fiscal years available")
else:
    selected_fys = st.sidebar.multiselect("Fiscal Year", options=all_fys)

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

    # Feature 3: Comparison mode toggle
    st.sidebar.markdown("**Visualization Options**")
    comparison_mode = st.sidebar.radio(
        "Department comparison:",
        options=["Aggregated", "Side-by-Side"],
        help="Aggregated: combines departments. Side-by-Side: separate bars for each department.",
        key="comparison_mode"
    )

    if view_mode == "Award View" and fa is not None:
        total_awd = pd.to_numeric(fa[AWARD_AMOUNT_COL], errors="coerce").sum()
        c1, c2 = st.columns(2)
        c1.metric("Award Count", f"{len(fa):,}")
        c2.metric("Total Obligated", f"${total_awd:,.0f}")

        if comparison_mode == "Side-by-Side" and selected_depts and len(selected_depts) > 1:
            awd_time = fa.groupby(["Fiscal Year", FACULTY_DEPT_COL, "Quarter"], as_index=False).agg(
                Count=("Fiscal Year", "size"), Dollars=(AWARD_AMOUNT_COL, "sum")
            )
            col1, col2 = st.columns(2)
            with col1:
                st.altair_chart(alt.Chart(awd_time).mark_bar().encode(
                    x="Fiscal Year:O", y="Count:Q", color=alt.Color("Quarter:O", sort=["1", "2", "3", "4"]),
                    column=alt.Column(f"{FACULTY_DEPT_COL}:N", title="Department"),
                    tooltip=["Fiscal Year", f"{FACULTY_DEPT_COL}", "Quarter", "Count"]
                ).properties(title="Award Count", height=350, width=180), use_container_width=True)
            with col2:
                st.altair_chart(alt.Chart(awd_time).mark_bar().encode(
                    x="Fiscal Year:O", y="Dollars:Q", color=alt.Color("Quarter:O", sort=["1", "2", "3", "4"]),
                    column=alt.Column(f"{FACULTY_DEPT_COL}:N", title="Department"),
                    tooltip=["Fiscal Year", f"{FACULTY_DEPT_COL}", "Quarter", alt.Tooltip("Dollars:Q", format="$,.0f")]
                ).properties(title="Award Dollars by Department", height=350, width=180), use_container_width=True)
        else:
            awd_time = fa.groupby(["Fiscal Year", "Quarter"], as_index=False).agg(
                Count=("Fiscal Year", "size"), Dollars=(AWARD_AMOUNT_COL, "sum")
            )
            col1, col2 = st.columns(2)
            with col1:
                st.altair_chart(alt.Chart(awd_time).mark_bar().encode(
                    x="Fiscal Year:O", y="Count:Q", color=alt.Color("Quarter:O", sort=["1", "2", "3", "4"]), tooltip=["Fiscal Year", "Quarter", "Count"]
                ).properties(title="Award Count", height=350), use_container_width=True)
            with col2:
                st.altair_chart(alt.Chart(awd_time).mark_bar().encode(
                    x="Fiscal Year:O", y="Dollars:Q", color=alt.Color("Quarter:O", sort=["1", "2", "3", "4"]),
                    tooltip=["Fiscal Year", "Quarter", alt.Tooltip("Dollars:Q", format="$,.0f")]
                ).properties(title="Award Dollars", height=350), use_container_width=True)

    elif view_mode == "Proposal View" and fp is not None:
        total_prop_val = pd.to_numeric(fp[PROPOSAL_AMOUNT_COL], errors="coerce").sum()
        # Feature 4: Success rate metric
        success_rate = None
        if PROPOSAL_FUNDED_FLAG_COL in fp.columns and len(fp):
            success_rate = normalize_funded_flag(fp[PROPOSAL_FUNDED_FLAG_COL]).mean()

        c1, c2, c3 = st.columns(3)
        c1.metric("Proposals Submitted", f"{len(fp):,}")
        c2.metric("Total Requested", f"${total_prop_val:,.0f}")
        c3.metric("Overall Success Rate", f"{success_rate:.1%}" if success_rate is not None else "—")

        if comparison_mode == "Side-by-Side" and selected_depts and len(selected_depts) > 1:
            prop_time = fp.groupby(["Fiscal Year", FACULTY_DEPT_COL, "Quarter"], as_index=False).agg(
                Count=("Fiscal Year", "size"), Dollars=(PROPOSAL_AMOUNT_COL, "sum")
            )
            pcol1, pcol2 = st.columns(2)
            with pcol1:
                st.altair_chart(alt.Chart(prop_time).mark_bar().encode(
                    x="Fiscal Year:O", y="Count:Q", color=alt.Color("Quarter:O", sort=["1", "2", "3", "4"]),
                    column=alt.Column(f"{FACULTY_DEPT_COL}:N", title="Department"),
                    tooltip=["Fiscal Year", f"{FACULTY_DEPT_COL}", "Quarter", "Count"]
                ).properties(title="Proposal Count", height=350, width=180), use_container_width=True)
            with pcol2:
                st.altair_chart(alt.Chart(prop_time).mark_bar().encode(
                    x="Fiscal Year:O", y="Dollars:Q", color=alt.Color("Quarter:O", sort=["1", "2", "3", "4"]),
                    column=alt.Column(f"{FACULTY_DEPT_COL}:N", title="Department"),
                    tooltip=["Fiscal Year", f"{FACULTY_DEPT_COL}", "Quarter", alt.Tooltip("Dollars:Q", format="$,.0f")]
                ).properties(title="Proposal Requested $", height=350, width=180), use_container_width=True)
        else:
            prop_time = fp.groupby(["Fiscal Year", "Quarter"], as_index=False).agg(
                Count=("Fiscal Year", "size"), Dollars=(PROPOSAL_AMOUNT_COL, "sum")
            )
            pcol1, pcol2 = st.columns(2)
            with pcol1:
                st.altair_chart(alt.Chart(prop_time).mark_bar().encode(
                    x="Fiscal Year:O", y="Count:Q", color=alt.Color("Quarter:O", sort=["1", "2", "3", "4"]), tooltip=["Fiscal Year", "Quarter", "Count"]
                ).properties(title="Proposal Count", height=350), use_container_width=True)
            with pcol2:
                st.altair_chart(alt.Chart(prop_time).mark_bar().encode(
                    x="Fiscal Year:O", y="Dollars:Q", color=alt.Color("Quarter:O", sort=["1", "2", "3", "4"]),
                    tooltip=["Fiscal Year", "Quarter", alt.Tooltip("Dollars:Q", format="$,.0f")]
                ).properties(title="Proposal Requested $", height=350), use_container_width=True)

    st.divider()
    st.markdown("### Breakdown Analysis")
    b1, b2 = st.columns(2)
    with b1:
        st.write("**By Department**")
        df_src = fa if view_mode == "Award View" else fp
        if df_src is not None:
            # Feature 4: Add success rates for proposals in the breakdown
            if view_mode == "Proposal View" and PROPOSAL_FUNDED_FLAG_COL in df_src.columns:
                dept_count = df_src.groupby(FACULTY_DEPT_COL).size().reset_index(name="Submitted")
                dept_funded = (
                    df_src[normalize_funded_flag(df_src[PROPOSAL_FUNDED_FLAG_COL])]
                      .groupby(FACULTY_DEPT_COL).size().reset_index(name="Funded")
                )
                dept_stats = dept_count.merge(dept_funded, on=FACULTY_DEPT_COL, how="left").fillna(0)
                dept_stats["Funded"] = dept_stats["Funded"].astype(int)
                dept_stats["Success Rate"] = (dept_stats["Funded"] / dept_stats["Submitted"]).map("{:.1%}".format)
            else:
                dept_stats = df_src.groupby(FACULTY_DEPT_COL).size().reset_index(name="Count")
            st.dataframe(dept_stats, hide_index=True, use_container_width=True)
    with b2:
        st.write("**By Sponsor (NIH Grouped)**")
        spon_col = AWARD_SPONSOR_COL if view_mode == "Award View" else PROP_SPONSOR_COL
        if df_src is not None:
            # Feature 4: Add success rates for sponsors in proposals view
            if view_mode == "Proposal View" and PROPOSAL_FUNDED_FLAG_COL in df_src.columns:
                spon_count = df_src.groupby(spon_col).size().reset_index(name="Submitted").sort_values("Submitted", ascending=False).head(10)
                spon_funded = (
                    df_src[normalize_funded_flag(df_src[PROPOSAL_FUNDED_FLAG_COL])]
                      .groupby(spon_col).size().reset_index(name="Funded")
                )
                spon_stats = spon_count.merge(spon_funded, on=spon_col, how="left").fillna(0)
                spon_stats["Funded"] = spon_stats["Funded"].astype(int)
                spon_stats["Success Rate"] = (spon_stats["Funded"] / spon_stats["Submitted"]).map("{:.1%}".format)
            else:
                spon_stats = df_src.groupby(spon_col).size().sort_values(ascending=False).head(10).reset_index(name="Count")
            st.dataframe(spon_stats, hide_index=True, use_container_width=True)

with tab_faculty:
    st.subheader("Faculty Drill-Down")

    ids_in_view = sorted(set(
        (fa[FACULTY_ID_COL].dropna().unique().tolist() if fa is not None else []) +
        (fp[FACULTY_ID_COL].dropna().unique().tolist() if fp is not None else [])
    ))

    if not ids_in_view:
        st.info("No faculty in the current filter selection.")
        st.stop()

    selected_faculty_info_drill = st.selectbox(
        "Select faculty",
        options=faculty_options,
        format_func=lambda x: x[1],
        key="faculty_drill_down_select"
    )
    selected_id = selected_faculty_info_drill[0]
    faculty_name = selected_faculty_info_drill[1]

    fa_pi = fa[fa[FACULTY_ID_COL] == selected_id] if fa is not None else None
    fp_pi = fp[fp[FACULTY_ID_COL] == selected_id] if fp is not None else None

    p_count = 0 if fp_pi is None else len(fp_pi)
    a_count = 0 if fa_pi is None else len(fa_pi)
    a_dollars = 0.0 if fa_pi is None else pd.to_numeric(fa_pi[AWARD_AMOUNT_COL], errors="coerce").fillna(0).sum()
    pi_success = None
    if fp_pi is not None and len(fp_pi):
        pi_success = normalize_funded_flag(fp_pi[PROPOSAL_FUNDED_FLAG_COL]).mean()

    st.markdown(f"#### {faculty_name} ({selected_id})")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total proposals", f"{p_count:,}")
    c2.metric("Total awards", f"{a_count:,}")
    c3.metric("Success rate", f"{pi_success:.1%}" if pi_success is not None else "—")
    c4.metric("Total award $", f"${a_dollars:,.0f}")

    st.divider()

    st.markdown("### Recent activity")
    r1, r2 = st.columns(2)
    with r1:
        if fp_pi is not None and len(fp_pi):
            st.write("Most recent proposal submission")
            display_cols_prop = [PROPOSAL_DATE_COL, "Proposal Project Title", "Proposal Lead Unit Name", PROP_SPONSOR_COL, "Fiscal Year"]
            cols_available = [c for c in display_cols_prop if c in fp_pi.columns]
            st.dataframe(fp_pi.sort_values(PROPOSAL_DATE_COL, ascending=False).head(5)[cols_available], use_container_width=True)
        else:
            st.caption("No proposals for this PI in current filters.")
    with r2:
        if fa_pi is not None and len(fa_pi):
            st.write("Most recent award")
            display_cols_award = [AWARD_DATE_COL, "Award Project Title", "Award Lead Unit Name", AWARD_SPONSOR_COL, AWARD_AMOUNT_COL, "Fiscal Year"]
            cols_available = [c for c in display_cols_award if c in fa_pi.columns]
            tmp_display = fa_pi.sort_values(AWARD_DATE_COL, ascending=False).head(5)[cols_available].copy()
            if AWARD_AMOUNT_COL in tmp_display.columns:
                tmp_display[AWARD_AMOUNT_COL] = tmp_display[AWARD_AMOUNT_COL].map("${:,.2f}".format)
            st.dataframe(tmp_display, use_container_width=True)
        else:
            st.caption("No awards for this PI in current filters.")

    st.divider()

    st.markdown("### Sponsor history")
    s1, s2 = st.columns(2)
    with s1:
        if fp_pi is not None and len(fp_pi):
            st.write("Top sponsors by submission count")
            st.dataframe(
                fp_pi.groupby(PROP_SPONSOR_COL).size().sort_values(ascending=False).head(10).reset_index(name="Submissions"),
                use_container_width=True,
                height=300
            )
    with s2:
        if fa_pi is not None and len(fa_pi):
            st.write("Top sponsors by award dollars")
            tmp = fa_pi.copy()
            tmp["_amt"] = pd.to_numeric(tmp[AWARD_AMOUNT_COL], errors="coerce").fillna(0)
            st.dataframe(
                tmp.groupby(AWARD_SPONSOR_COL)["_amt"].sum().sort_values(ascending=False).head(10).reset_index(name="Award $"),
                use_container_width=True,
                height=300
            )

with tab_tables:
    st.subheader("Tables (filtered, full datasets)")

    t1, t2 = st.columns(2)

    with t1:
        st.markdown("### Proposals (filtered)")
        if fp is not None:
            st.dataframe(fp, use_container_width=True, height=550)
            st.download_button(
                "Download Filtered Proposals (Excel)",
                data=to_excel(fp) if fp is not None else b"",
                file_name="Filtered_Proposals.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.caption("No proposals uploaded.")

    with t2:
        st.markdown("### Awards (filtered)")
        if fa is not None:
            st.dataframe(fa, use_container_width=True, height=550)
            st.download_button(
                "Download Filtered Awards (Excel)",
                data=to_excel(fa) if fa is not None else b"",
                file_name="Filtered_Awards.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.caption("No awards uploaded.")


def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    return output.getvalue()
