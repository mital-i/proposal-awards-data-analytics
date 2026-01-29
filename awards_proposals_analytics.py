import streamlit as st
import pandas as pd
from io import BytesIO

def load_data(uploaded_file):
    try:
        return pd.read_excel(uploaded_file, sheet_name=0)
    except Exception:
        uploaded_file.seek(0)
        try:
            return pd.read_csv(
                uploaded_file, 
                sep=None, 
                engine='python', 
                on_bad_lines='skip'
            )
        except Exception as e:
            st.error(f"Could not parse file: {e}")
            return None

st.set_page_config(page_title="BioSci Awards & Proposals", layout="wide")
st.title("📊 Research Development Data Processor")

st.sidebar.header("Configuration")
master_file = st.sidebar.file_uploader(
    "Upload Faculty_Master File", 
    type=['xlsx', 'xls', 'csv']
)

def get_fiscal_quarter(df, date_col):
    """Calculates Fiscal Year and Quarter based on July 1st start."""
    df[date_col] = pd.to_datetime(df[date_col])
    df['Fiscal Year'] = df[date_col].dt.year + (df[date_col].dt.month >= 7).astype(int)
    df['Quarter'] = ((df[date_col].dt.month - 7) % 12 // 3) + 1
    return df

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

tab1, tab2 = st.tabs(["Awards", "Proposals"])

if master_file:
    # Use the robust loader that ignores sheet names
    faculty_master = load_data(master_file)
    
    # Standardize column names to match R script expectations
    faculty_master['Award PI Campus ID'] = pd.to_numeric(faculty_master['Award PI Campus ID'], errors='coerce')
    depts_and_ids = faculty_master[["Award PI Campus ID", "Department"]]
    campus_ids = depts_and_ids["Award PI Campus ID"].tolist()

    with tab1:
        st.header("Awards Processing")
        award_file = st.file_uploader("Upload awards_df", type=['xlsx', 'csv', 'xls'])
        
        if award_file:
            dwq_awards = load_data(award_file)
            
            # Filter BioSci Only
            biosci_awards = dwq_awards[dwq_awards['Award PI Campus ID'].isin(campus_ids)]
            biosci_awards = biosci_awards.merge(depts_and_ids, on='Award PI Campus ID', how='left')
            
            # Apply Fiscal Logic and filter for 2025/2026
            final_awards = get_fiscal_quarter(biosci_awards, 'Award Finalize Date')
            
            st.write("### Preview: Processed Awards")
            st.dataframe(final_awards.head(10))
            
            st.download_button(
                label="Download Processed Awards",
                data=to_excel(final_awards),
                file_name="Processed_Awards.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    with tab2:
        st.header("Proposals Processing")
        proposal_file = st.file_uploader("Upload proposals_df", type=['xlsx', 'csv', 'xls'])
        
        if proposal_file:
            dwq_proposals = load_data(proposal_file)
            # Rename PI ID column to match Master ID for merging
            dwq_proposals = dwq_proposals.rename(columns={'Proposal PI Campus ID': 'Award PI Campus ID'})
            
            # Filter BioSci Only
            biosci_proposals = dwq_proposals[dwq_proposals['Award PI Campus ID'].isin(campus_ids)]
            biosci_proposals = biosci_proposals.merge(depts_and_ids, on='Award PI Campus ID', how='left')
            
            # Apply Fiscal Logic
            final_proposals = get_fiscal_quarter(biosci_proposals, 'Proposal Process Date')
            
            st.write("### Preview: Processed Proposals")
            st.dataframe(final_proposals.head(10))
            
            st.download_button(
                label="Download Processed Proposals",
                data=to_excel(final_proposals),
                file_name="Processed_Proposals.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

else:
    st.info("Please upload your Faculty Master file in the sidebar to begin.")