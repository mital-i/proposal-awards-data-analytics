import streamlit as st
import pandas as pd
import altair as alt
from io import BytesIO

FACULTY_MASTER_PATH = "data/Faculty_Master.xlsx"

AWARD_DATE_COL = "Award Finalize Date"
AWARD_AMOUNT_COL = "Award Obligated Total Cost"
AWARD_SPONSOR_COL = "Award Sponsor Name"
AWARD_TRANS_TYPE_COL = "Award Transaction Type Description"
NIH_ACTIVITY_CODE_COL = "NIH Activity Code"

PROPOSAL_DATE_COL = "Proposal Process Date"
PROPOSAL_FUNDED_FLAG_COL = "Proposal Funded Flag"
PROP_SPONSOR_COL = "Proposal Sponsor Name"

FACULTY_ID_COL = "Award PI Campus ID"
FACULTY_DEPT_COL = "Department"


def load_data(uploaded_file):
    try:
        return pd.read_excel(uploaded_file, sheet_name=0)
    except Exception:
        uploaded_file.seek(0)
        try:
            return pd.read_csv(uploaded_file, sep=None, engine="python", on_bad_lines="skip")
        except Exception as e:
            st.error(f"Could not parse file: {e}")
            return None


@st.cache_data
def load_faculty_master():
    df = pd.read_excel(FACULTY_MASTER_PATH)
    df[FACULTY_ID_COL] = pd.to_numeric(df[FACULTY_ID_COL], errors="coerce")
    return df


def get_fiscal_quarter(df, date_col):
    df = df.copy()
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    df = df.dropna(subset=[date_col])
    df["Fiscal Year"] = df[date_col].dt.year + (df[date_col].dt.month >= 7).astype(int)
    df["Quarter"] = ((df[date_col].dt.month - 7) % 12 // 3) + 1
    df["Fiscal Quarter"] = "Q" + df["Quarter"].astype(str)
    return df


def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    return output.getvalue()


def normalize_funded_flag(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip().str.lower()
    return s.isin(["y", "yes", "true", "1", "funded"])

st.set_page_config(page_title="BioSci Awards & Proposals", layout="wide")
st.title("Research Development Portfolio Tool")

try:
    faculty_master = load_faculty_master()
except Exception as e:
    st.error(f"Could not load Faculty_Master from '{FACULTY_MASTER_PATH}'. Error: {e}")
    st.stop()

depts_and_ids = faculty_master[[FACULTY_ID_COL, FACULTY_DEPT_COL]].drop_duplicates()
campus_ids = depts_and_ids[FACULTY_ID_COL].dropna().tolist()

st.sidebar.header("Uploads")
award_file = st.sidebar.file_uploader("Upload awards_df", type=["xlsx", "csv", "xls"])
proposal_file = st.sidebar.file_uploader("Upload proposals_df", type=["xlsx", "csv", "xls"])

final_awards = None
final_proposals = None

if award_file:
    dwq_awards = load_data(award_file)
    if dwq_awards is not None and FACULTY_ID_COL in dwq_awards.columns:
        biosci_awards = dwq_awards[dwq_awards[FACULTY_ID_COL].isin(campus_ids)].copy()
        biosci_awards = biosci_awards.merge(depts_and_ids, on=FACULTY_ID_COL, how="left")
        final_awards = get_fiscal_quarter(biosci_awards, AWARD_DATE_COL)

if proposal_file:
    dwq_proposals = load_data(proposal_file)
    if dwq_proposals is not None:
        if "Proposal PI Campus ID" in dwq_proposals.columns and FACULTY_ID_COL not in dwq_proposals.columns:
            dwq_proposals = dwq_proposals.rename(columns={"Proposal PI Campus ID": FACULTY_ID_COL})

        if FACULTY_ID_COL in dwq_proposals.columns:
            biosci_proposals = dwq_proposals[dwq_proposals[FACULTY_ID_COL].isin(campus_ids)].copy()
            biosci_proposals = biosci_proposals.merge(depts_and_ids, on=FACULTY_ID_COL, how="left")
            final_proposals = get_fiscal_quarter(biosci_proposals, PROPOSAL_DATE_COL)

if final_awards is None and final_proposals is None:
    st.info("Upload awards and/or proposals in the sidebar to begin.")
    st.stop()

st.sidebar.header("Dashboard View")
view_mode = st.sidebar.radio("Select View", options=["Award View", "Proposal View"])

st.sidebar.header("Filters")

fy_values = sorted(set(
    (final_awards["Fiscal Year"].unique().tolist() if final_awards is not None else []) +
    (final_proposals["Fiscal Year"].unique().tolist() if final_proposals is not None else [])
))
selected_fys = st.sidebar.multiselect("Fiscal Year", options=fy_values) 

dept_values = sorted(set(
    (final_awards[FACULTY_DEPT_COL].dropna().unique().tolist() if final_awards is not None else []) +
    (final_proposals[FACULTY_DEPT_COL].dropna().unique().tolist() if final_proposals is not None else [])
))
selected_depts = st.sidebar.multiselect("Department", options=dept_values) 

nih_values = sorted(set(
    (final_awards[NIH_ACTIVITY_CODE_COL].dropna().unique().tolist() if final_awards is not None else []) +
    (final_proposals[NIH_ACTIVITY_CODE_COL].dropna().unique().tolist() if final_proposals is not None else [])
))
selected_nih = st.sidebar.multiselect("NIH Activity Code", options=nih_values)

faculty_names_map = faculty_master.set_index(FACULTY_ID_COL)["Name"].to_dict()
faculty_options = sorted([(fid, faculty_names_map.get(fid, f"ID: {fid}")) for fid in campus_ids], key=lambda x: x[1])

selected_faculty_info = st.sidebar.multiselect(
    "Faculty",
    options=faculty_options,
    format_func=lambda x: x[1]
)
selected_faculty_ids = [x[0] for x in selected_faculty_info]

def apply_filters(df, is_award=False):
    if df is None:
        return None
    out = df.copy()

    if is_award and AWARD_TRANS_TYPE_COL in out.columns:
        out = out[out[AWARD_TRANS_TYPE_COL].isin(["New", "Renewal", "Supplement"])]

    if selected_fys:
        out = out[out["Fiscal Year"].isin(selected_fys)]
    if selected_depts:
        out = out[out[FACULTY_DEPT_COL].isin(selected_depts)]
    if selected_nih:
        if NIH_ACTIVITY_CODE_COL in out.columns:
            out = out[out[NIH_ACTIVITY_CODE_COL].isin(selected_nih)]
    if selected_faculty_ids:
        out = out[out[FACULTY_ID_COL].isin(selected_faculty_ids)]

    return out

fa = apply_filters(final_awards, is_award=True)
fp = apply_filters(final_proposals)

tab_overview, tab_faculty, tab_tables = st.tabs(["Portfolio Overview", "Faculty Drill-Down", "Tables & Downloads"])

with tab_overview:
    st.subheader(f"Portfolio Overview - {view_mode}")

    if view_mode == "Award View":
        total_awards = 0 if fa is None else len(fa)
        award_dollars = 0.0
        if fa is not None:
            award_dollars = pd.to_numeric(fa[AWARD_AMOUNT_COL], errors="coerce").fillna(0).sum()
        
        c1, c2 = st.columns(2)
        c1.metric("Total awards (count)", f"{total_awards:,}")
        c2.metric("Total awards ($)", f"${award_dollars:,.0f}")
    else:
        total_proposals = 0 if fp is None else len(fp)
        success_rate = None
        if fp is not None and PROPOSAL_FUNDED_FLAG_COL in fp.columns and len(fp):
            success_rate = normalize_funded_flag(fp[PROPOSAL_FUNDED_FLAG_COL]).mean()
        
        c1, c2 = st.columns(2)
        c1.metric("Total proposals submitted", f"{total_proposals:,}")
        c2.metric("Overall success rate", f"{success_rate:.1%}" if success_rate is not None else "—")

    st.divider()

    if view_mode == "Proposal View" and fp is not None and len(fp):
        st.markdown("#### Proposals submitted by Fiscal Year / Quarter")
        prop_time = fp.groupby(["Fiscal Year", "Quarter"], as_index=False).size().rename(columns={"size": "Proposals"})
        
        chart_prop = alt.Chart(prop_time).mark_bar().encode(
            x=alt.X("Fiscal Year:O"),
            y=alt.Y("Proposals:Q"),
            color=alt.Color("Quarter:O"),
            tooltip=["Fiscal Year", "Quarter", "Proposals"]
        ).properties(height=300)

        if selected_fys and len(selected_fys) == 1 and selected_depts and len(selected_depts) > 1:
            st.markdown(f"**Side-by-Side Comparison by Department for FY {selected_fys[0]}**")
            dept_comp = fp.groupby([FACULTY_DEPT_COL, "Quarter"], as_index=False).size().rename(columns={"size": "Proposals"})
            chart_dept = alt.Chart(dept_comp).mark_bar().encode(
                x=alt.X(f"{FACULTY_DEPT_COL}:N", title="Department"),
                y=alt.Y("Proposals:Q"),
                color=alt.Color("Quarter:O"),
                tooltip=[FACULTY_DEPT_COL, "Quarter", "Proposals"]
            ).properties(height=300)
            st.altair_chart(chart_dept, use_container_width=True)
        else:
            st.altair_chart(chart_prop, use_container_width=True)

    if view_mode == "Award View" and fa is not None and len(fa):
        st.markdown("#### Awards received by Fiscal Year (count and $)")
        awards_count = fa.groupby("Fiscal Year", as_index=False).size().rename(columns={"size": "Awards"})
        awards_dollars = (
            fa.assign(_amt=pd.to_numeric(fa[AWARD_AMOUNT_COL], errors="coerce").fillna(0))
              .groupby("Fiscal Year", as_index=False)["_amt"].sum()
              .rename(columns={"_amt": "Award Dollars"})
        )
        awards_group = awards_count.merge(awards_dollars, on="Fiscal Year", how="left")

        col1, col2 = st.columns(2)
        with col1:
            st.altair_chart(
                alt.Chart(awards_group).mark_bar().encode(
                    x=alt.X("Fiscal Year:O"),
                    y=alt.Y("Awards:Q"),
                    tooltip=["Fiscal Year", "Awards"]
                ).properties(height=300, title="Award Count by Fiscal Year"),
                use_container_width=True
            )
        with col2:
            st.altair_chart(
                alt.Chart(awards_group).mark_bar().encode(
                    x=alt.X("Fiscal Year:O"),
                    y=alt.Y("Award Dollars:Q"),
                    tooltip=[
                        alt.Tooltip("Fiscal Year:O"),
                        alt.Tooltip("Award Dollars:Q", format="$,.0f")
                    ]
                ).properties(height=300, title="Award $ by Fiscal Year"),
                use_container_width=True
            )
        
        if selected_fys and len(selected_fys) == 1 and selected_depts and len(selected_depts) > 1:
            st.markdown(f"**Side-by-Side Comparison by Department for FY {selected_fys[0]}**")
            dept_comp = (
                fa.assign(_amt=pd.to_numeric(fa[AWARD_AMOUNT_COL], errors="coerce").fillna(0))
                  .groupby(FACULTY_DEPT_COL, as_index=False)["_amt"].sum()
                  .rename(columns={"_amt": "Award Dollars"})
            )
            chart_dept = alt.Chart(dept_comp).mark_bar().encode(
                x=alt.X(f"{FACULTY_DEPT_COL}:N", title="Department"),
                y=alt.Y("Award Dollars:Q"),
                tooltip=[
                    alt.Tooltip(f"{FACULTY_DEPT_COL}:N"),
                    alt.Tooltip("Award Dollars:Q", format="$,.0f")
                ]
            ).properties(height=300, title="Award Dollars by Department")
            st.altair_chart(chart_dept, use_container_width=True)

    st.divider()

    st.markdown("#### Breakdowns")

    b1, b2 = st.columns(2)
    with b1:
        st.markdown("**By Department**")
        if view_mode == "Proposal View":
            dept_summary = (fp.groupby(FACULTY_DEPT_COL).size().reset_index(name="Proposals") if fp is not None else pd.DataFrame(columns=[FACULTY_DEPT_COL, "Proposals"]))
        else:
            dept_counts = (fa.groupby(FACULTY_DEPT_COL).size().rename("Awards") if fa is not None else pd.Series(dtype=int))
            dept_dollars = (
                fa.assign(_amt=pd.to_numeric(fa[AWARD_AMOUNT_COL], errors="coerce").fillna(0))
                  .groupby(FACULTY_DEPT_COL)["_amt"].sum().rename("Award Dollars")
                if fa is not None else pd.Series(dtype=float)
            )
            dept_summary = pd.concat([dept_counts, dept_dollars], axis=1).fillna(0).reset_index()
            if "Award Dollars" in dept_summary.columns:
                dept_summary["Award Dollars"] = dept_summary["Award Dollars"].map("${:,.0f}".format)

        st.dataframe(dept_summary, use_container_width=True, height=350)

    with b2:
        if view_mode == "Proposal View":
            st.markdown("**By Sponsor (Proposals)**")
            if fp is not None and PROP_SPONSOR_COL in fp.columns:
                top_prop_sponsors = (
                    fp.groupby(PROP_SPONSOR_COL).size()
                      .sort_values(ascending=False).head(20)
                      .reset_index(name="Proposal submissions")
                )
                st.dataframe(top_prop_sponsors, use_container_width=True, height=350)
            else:
                st.caption("No proposals uploaded or sponsor column missing.")
        else:
            st.markdown("**By Sponsor (Awards)**")
            if fa is not None and AWARD_SPONSOR_COL in fa.columns:
                tmp = fa.copy()
                tmp["_amt"] = pd.to_numeric(tmp[AWARD_AMOUNT_COL], errors="coerce").fillna(0)
                top_awd_sponsors = (
                    tmp.groupby(AWARD_SPONSOR_COL)["_amt"].sum()
                      .sort_values(ascending=False).head(20)
                      .reset_index(name="Award Dollars")
                )
                top_awd_sponsors["Award Dollars"] = top_awd_sponsors["Award Dollars"].map("${:,.0f}".format)
                st.dataframe(top_awd_sponsors, use_container_width=True, height=350)
            else:
                st.caption("No awards uploaded or sponsor column missing.")

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
                data=to_excel(fp),
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
                data=to_excel(fa),
                file_name="Filtered_Awards.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.caption("No awards uploaded.")