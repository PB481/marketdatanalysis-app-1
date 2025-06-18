import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt
import seaborn as sns

st.set_page_config(layout="wide")

st.title("Excel Data Analyzer: Fund Insights")

st.write(
    """
    Upload multiple Excel files containing fund data. The app will extract common headers,
    specifically 'Domicile' and 'Legal Status', and provide an analysis of these fields
    across all uploaded files.
    """
)

# Define the common headers we are looking for (case-insensitive for robustness)
COMMON_HEADERS = [
    "Domicile", "Legal Status", "Promoter/Initiator", "Monterey SchemelD",
    "Fund Name", "Sub-Fund Name", "Region & Category", "Industry",
    "Asset Allocation", "Fund Of Funds / Fund of Hedge Funds", "Master/Feeder",
    "Monterey Admin ID", "Administrator Location", "Monterey Audit ID",
    "Auditor", "Auditor Location", "Monterey Leg ID", "Legal Adviser",
    "Legal Adviser Location", "Transfer Agent", "Monterey ManCo/AIFM ID",
    "ManCo/AlFM Location", "ManCo/AIFM Parent Origin", "ManCo/AIFM Third Party",
    "Registered AIFM", "Self Managed", "Fund Vintage Year/Launch Date",
    "Sub-Fund Vintage Year/Launch Date", "Promoter Origin Code", "Administrator",
    "UCITS/ AIF", "TNAV USD", "USS TNAV"
]

# Create a placeholder for the dataframe that will hold all consolidated data
all_data_df = pd.DataFrame()
processed_files_count = 0
found_headers_across_files = set()

uploaded_files = st.file_uploader(
    "Upload your Excel files", type=["xlsx", "xls"], accept_multiple_files=True
)

if uploaded_files:
    st.subheader("Processing Files...")
    progress_bar = st.progress(0)
    for i, uploaded_file in enumerate(uploaded_files):
        try:
            # Read the Excel file
            df = pd.read_excel(uploaded_file)

            st.write(f"--- Processing: **{uploaded_file.name}** ---")
            st.write(f"Original columns in '{uploaded_file.name}': {', '.join(df.columns.tolist())}")

            # Normalize column names for robust matching (e.g., remove leading/trailing spaces, lowercase)
            df.columns = df.columns.str.strip()
            df.columns = df.columns.str.replace('.', '', regex=False) # Remove periods if present

            # Identify common headers present in the current file
            current_file_headers = []
            for header in COMMON_HEADERS:
                # Case-insensitive matching
                matching_cols = [col for col in df.columns if col.lower() == header.lower()]
                if matching_cols:
                    current_file_headers.append(matching_cols[0]) # Take the actual column name from df
                    found_headers_across_files.add(matching_cols[0])

            if not current_file_headers:
                st.warning(f"No common headers found in '{uploaded_file.name}'. Skipping this file.")
                continue

            # Extract relevant data: all common headers found, plus 'Domicile' and 'Legal Status' specifically
            # Ensure 'Domicile' and 'Legal Status' are always sought first
            cols_to_extract = []
            for col_name in ['Domicile', 'Legal Status'] + [h for h in current_file_headers if h not in ['Domicile', 'Legal Status']]:
                matching_cols = [col for col in df.columns if col.lower() == col_name.lower()]
                if matching_cols:
                    cols_to_extract.append(matching_cols[0])

            extracted_df = df[cols_to_extract].copy()
            extracted_df['Source File'] = uploaded_file.name
            all_data_df = pd.concat([all_data_df, extracted_df], ignore_index=True)
            processed_files_count += 1

            st.success(f"Successfully processed '{uploaded_file.name}'.")
            st.write(f"Extracted headers from this file: {', '.join(cols_to_extract)}")

        except Exception as e:
            st.error(f"Error processing {uploaded_file.name}: {e}")
        finally:
            progress_bar.progress((i + 1) / len(uploaded_files))

    if not all_data_df.empty:
        st.success(f"Successfully processed {processed_files_count} file(s) and consolidated data.")
        st.subheader("Consolidated Data Preview")
        st.dataframe(all_data_df.head())

        st.subheader("Detailed Analysis & Insights")

        # --- Analysis of Domicile ---
        st.markdown("### Domicile Analysis")
        if 'Domicile' in all_data_df.columns:
            st.write("Distribution of Domicile:")
            domicile_counts = all_data_df['Domicile'].value_counts(dropna=False).sort_values(ascending=False)
            st.dataframe(domicile_counts)

            fig1, ax1 = plt.subplots(figsize=(10, 6))
            sns.barplot(x=domicile_counts.index, y=domicile_counts.values, ax=ax1, palette='viridis')
            ax1.set_title('Domicile Distribution')
            ax1.set_xlabel('Domicile')
            ax1.set_ylabel('Count')
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            st.pyplot(fig1)

            st.write("Top 5 Domiciles:")
            st.write(domicile_counts.head(5))

            st.write(f"Number of unique Domiciles: **{all_data_df['Domicile'].nunique()}**")
        else:
            st.info("No 'Domicile' column found in the consolidated data.")

        # --- Analysis of Legal Status ---
        st.markdown("### Legal Status Analysis")
        if 'Legal Status' in all_data_df.columns:
            st.write("Distribution of Legal Status:")
            legal_status_counts = all_data_df['Legal Status'].value_counts(dropna=False).sort_values(ascending=False)
            st.dataframe(legal_status_counts)

            fig2, ax2 = plt.subplots(figsize=(10, 6))
            sns.barplot(x=legal_status_counts.index, y=legal_status_counts.values, ax=ax2, palette='magma')
            ax2.set_title('Legal Status Distribution')
            ax2.set_xlabel('Legal Status')
            ax2.set_ylabel('Count')
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            st.pyplot(fig2)

            st.write("Top 5 Legal Status types:")
            st.write(legal_status_counts.head(5))

            st.write(f"Number of unique Legal Status types: **{all_data_df['Legal Status'].nunique()}**")
        else:
            st.info("No 'Legal Status' column found in the consolidated data.")

        # --- Common Themes and Insights (General) ---
        st.markdown("### Common Themes and General Insights")

        if 'Fund Name' in all_data_df.columns:
            st.write(f"Total unique 'Fund Name' entries: **{all_data_df['Fund Name'].nunique()}**")
            st.write("Most frequent Fund Names:")
            st.write(all_data_df['Fund Name'].value_counts().head())

        if 'Promoter/Initiator' in all_data_df.columns:
            st.write(f"Total unique 'Promoter/Initiator' entries: **{all_data_df['Promoter/Initiator'].nunique()}**")
            st.write("Most frequent Promoters/Initiators:")
            st.write(all_data_df['Promoter/Initiator'].value_counts().head())

        if 'Industry' in all_data_df.columns:
            st.write(f"Total unique 'Industry' entries: **{all_data_df['Industry'].nunique()}**")
            st.write("Distribution of Industry:")
            industry_counts = all_data_df['Industry'].value_counts(dropna=False).sort_values(ascending=False)
            st.dataframe(industry_counts)
            fig3, ax3 = plt.subplots(figsize=(12, 7))
            sns.countplot(y='Industry', data=all_data_df, order=all_data_df['Industry'].value_counts().index, palette='crest', ax=ax3)
            ax3.set_title('Industry Distribution')
            ax3.set_xlabel('Count')
            ax3.set_ylabel('Industry')
            plt.tight_layout()
            st.pyplot(fig3)

        if 'Asset Allocation' in all_data_df.columns:
            st.write(f"Total unique 'Asset Allocation' entries: **{all_data_df['Asset Allocation'].nunique()}**")
            st.write("Distribution of Asset Allocation:")
            asset_allocation_counts = all_data_df['Asset Allocation'].value_counts(dropna=False).sort_values(ascending=False)
            st.dataframe(asset_allocation_counts)
            fig4, ax4 = plt.subplots(figsize=(12, 7))
            sns.countplot(y='Asset Allocation', data=all_data_df, order=all_data_df['Asset Allocation'].value_counts().index, palette='plasma', ax=ax4)
            ax4.set_title('Asset Allocation Distribution')
            ax4.set_xlabel('Count')
            ax4.set_ylabel('Asset Allocation')
            plt.tight_layout()
            st.pyplot(fig4)


        st.subheader("Summary of Found Common Headers Across All Files")
        if found_headers_across_files:
            st.write(f"The following common headers were found and extracted from at least one file: {', '.join(sorted(list(found_headers_across_files)))}")
        else:
            st.warning("No common headers were successfully extracted from any of the uploaded files.")

    elif uploaded_files: # Means files were uploaded but none were processed successfully
        st.warning("No data could be processed from the uploaded files. Please check file formats and column headers.")

else:
    st.info("Please upload your Excel files to begin the analysis.")
