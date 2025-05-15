import streamlit as st
import pandas as pd
from db_utils import sync_sheet_to_db

st.set_page_config(page_title="Excel to SQL Sync App", layout="wide")
st.title("📊 Excel to SQL Server Synchronization System")

uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names

        st.success(f"✅ Found {len(sheet_names)} sheet(s).")
        selected_sheets = st.multiselect("Select sheets to sync:", sheet_names)

        if selected_sheets:
            for sheet in selected_sheets:
                st.subheader(f"📄 Sheet: `{sheet}`")
                df = xl.parse(sheet)

                if "id" not in df.columns:
                    st.error("❌ Missing required primary key column: `id`")
                    continue

                st.dataframe(df)

                if st.button(f"Sync '{sheet}' to Database", key=sheet):
                    with st.spinner("🔄 Synchronizing..."):
                        try:
                            inserted, updated = sync_sheet_to_db(df)
                            st.success(f"✅ Sheet '{sheet}' synced successfully! Inserted: {inserted}, Updated: {updated}")
                        except Exception as e:
                            st.error(f"⚠️ Error syncing sheet '{sheet}': {e}")
    except Exception as e:
        st.error(f"❌ Failed to read Excel file: {e}")
