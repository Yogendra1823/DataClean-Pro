import streamlit as st
import pandas as pd
import io
import plotly.express as px
from clean_and_report import (
    inspect_data, fix_column_names, convert_numeric_columns, remove_duplicates,
    clean_text_columns, standardise_categories, fill_missing_values,
    parse_dates, flag_outliers, add_derived_columns, generate_excel_report
)

st.set_page_config(page_title="DataClean Pro", page_icon="✨", layout="wide")

# -- CSS overrides for an even cleaner look --
st.markdown("""
    <style>
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    h1 { color: #0078D4; }
    </style>
""", unsafe_allow_html=True)

st.title("✨ DataClean Pro Dashboard")
st.markdown("Transform messy datasets into pristine data and visual reports instantly.")

# -- 1. FILE UPLOAD --
with st.container():
    uploaded_file = st.file_uploader("Drop your CSV or Excel file here", type=["csv", "xlsx", "xls"])

if not uploaded_file:
    st.info("👆 Upload a file to get started.")
    st.stop()

# -- 2. LOAD DATA --
try:
    if uploaded_file.name.endswith(".csv"):
        df_raw = pd.read_csv(uploaded_file)
    else:
        df_raw = pd.read_excel(uploaded_file)
except Exception as e:
    st.error(f"Error loading file: {e}")
    st.stop()

df = fix_column_names(df_raw.copy())
all_columns = list(df.columns)

# -- 3. CONFIGURATION SIDEBAR --
with st.sidebar:
    st.header("⚙️ Column Configuration")
    st.markdown("Assign columns to their correct types for optimal cleaning.")
    
    numeric_cols = st.multiselect("Numeric Columns (fill missing with median)", all_columns)
    text_cols = st.multiselect("Text Columns (fix casing & spaces)", all_columns)
    date_cols = st.multiselect("Date Columns (standardize format)", all_columns)
    
    st.markdown("---")
    st.subheader("Chart Settings")
    category_col = st.selectbox("Category Column (for X-axis / Grouping)", ["(None)"] + all_columns)
    value_col = st.selectbox("Value Column (for Y-axis / Measurement)", ["(None)"] + all_columns)
    
    run_button = st.button("🚀 Clean Data & Generate Report", use_container_width=True, type="primary")

# -- 4. MAIN WORKSPACE TABS --
tab_raw, tab_clean, tab_visuals = st.tabs(["📄 Raw Data", "✨ Cleaned Data", "📊 Visual Summaries"])

with tab_raw:
    st.subheader("Raw Data Preview")
    st.dataframe(df_raw.head(20), use_container_width=True)
    st.caption(f"Total Rows: {len(df_raw)} | Total Columns: {len(df_raw.columns)}")

if run_button:
    with st.spinner("Applying enterprise-grade data cleaning..."):
        df_original = df.copy()
        problems_before = inspect_data(df)
        
        # Apply pipeline
        df = convert_numeric_columns(df, numeric_cols)
        df = remove_duplicates(df)
        df = clean_text_columns(df, text_cols)
        if category_col != "(None)":
            df = standardise_categories(df, category_col)
        df = fill_missing_values(df, numeric_cols, text_cols)
        df = parse_dates(df, date_cols)
        if len(date_cols) > 0:
            df = add_derived_columns(df, date_col=date_cols[0])
        if value_col != "(None)":
            df = flag_outliers(df, value_col)
            
        rows_removed = len(df_original) - len(df)
        blanks_fixed = sum(v for k, v in problems_before.items() if "Missing" in k and isinstance(v, int))
        
        # Store in session state so it persists across tab switches
        st.session_state['cleaned'] = True
        st.session_state['df_cleaned'] = df
        st.session_state['df_original'] = df_original
        st.session_state['problems'] = problems_before
        st.session_state['stats'] = (rows_removed, blanks_fixed)

if st.session_state.get('cleaned', False):
    df_cleaned = st.session_state['df_cleaned']
    problems = st.session_state['problems']
    rows_removed, blanks_fixed = st.session_state['stats']
    
    with tab_clean:
        # KPI Metrics
        st.subheader("Data Quality Improvements")
        col1, col2, col3 = st.columns(3)
        col1.metric("Rows Processed", len(df_cleaned), f"-{rows_removed} Duplicates" if rows_removed > 0 else "")
        col2.metric("Missing Values Fixed", blanks_fixed, f"+{blanks_fixed} Replaced", delta_color="normal")
        col3.metric("Issues Detected", len(problems), "Resolved", delta_color="inverse")
        
        st.markdown("---")
        st.subheader("Cleaned Data Preview")
        st.dataframe(df_cleaned.head(30), use_container_width=True)
        
        # Download Section
        st.markdown("### 📥 Download Outputs")
        csv_buffer = io.StringIO()
        df_cleaned.to_csv(csv_buffer, index=False)
        
        import tempfile
        import os
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            temp_excel_path = tmp.name
            
        generate_excel_report(st.session_state['df_original'], df_cleaned, problems, temp_excel_path)
        with open(temp_excel_path, "rb") as f:
            excel_data = f.read()
        os.remove(temp_excel_path)
        
        dl_col1, dl_col2 = st.columns(2)
        dl_col1.download_button("Download Cleaned CSV", data=csv_buffer.getvalue(), file_name="cleaned_data.csv", mime="text/csv", use_container_width=True)
        dl_col2.download_button("Download Excel Report", data=excel_data, file_name="cleaning_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    with tab_visuals:
        if category_col != "(None)" and value_col != "(None)" and category_col in df_cleaned.columns and value_col in df_cleaned.columns:
            st.subheader(f"Total {value_col} by {category_col}")
            
            # Aggregate data for charting
            chart_df = df_cleaned.groupby(category_col)[value_col].sum().reset_index()
            chart_df = chart_df.sort_values(by=value_col, ascending=False)
            
            # Plotly Interactive Chart
            fig = px.bar(chart_df, x=category_col, y=value_col, text_auto='.2s', 
                         color=value_col, color_continuous_scale="Blues",
                         title=f"Distribution of {value_col}")
            fig.update_layout(xaxis_title=category_col, yaxis_title=f"Total {value_col}", template="plotly_dark")
            
            st.plotly_chart(fig, use_container_width=True)
            
            # Data Distribution Pie Chart
            fig_pie = px.pie(chart_df, names=category_col, values=value_col, title=f"Proportion of {value_col}")
            fig_pie.update_layout(template="plotly_dark")
            st.plotly_chart(fig_pie, use_container_width=True)
            
        else:
            st.info("Configure the Category and Value columns in the sidebar and run the cleaning process to generate visual summaries.")
