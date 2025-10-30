import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Spreadsheet Comparison Tool", page_icon="📊", layout="wide")

st.title("📊 Spreadsheet Comparison Tool")
st.markdown("Compare two spreadsheet files and identify deleted, added, and modified rows.")

# Initialize session state
if 'layout_confirmed' not in st.session_state:
    st.session_state.layout_confirmed = False
if 'df_before' not in st.session_state:
    st.session_state.df_before = None
if 'df_after' not in st.session_state:
    st.session_state.df_after = None

# File uploaders
col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 Before File")
    file_before = st.file_uploader("Upload the 'before' file", type=['csv', 'xls', 'xlsx'], key='before')

with col2:
    st.subheader("📁 After File")
    file_after = st.file_uploader("Upload the 'after' file", type=['csv', 'xls', 'xlsx'], key='after')


def read_file(file):
    """Read various spreadsheet formats"""
    if file.name.endswith('.csv'):
        return pd.read_csv(file)
    elif file.name.endswith(('.xls', '.xlsx')):
        return pd.read_excel(file)
    return None


def compare_dataframes(df_before, df_after, key_columns, compare_columns):
    """Compare two dataframes and identify differences"""
    # Create composite keys from multiple identifier columns
    df_before['_key'] = df_before[key_columns].astype(str).agg('||'.join, axis=1)
    df_after['_key'] = df_after[key_columns].astype(str).agg('||'.join, axis=1)

    # Select only key columns and compare columns for analysis
    cols_to_use = key_columns + compare_columns

    # Find deleted rows (in before but not in after)
    deleted_keys = df_before[~df_before['_key'].isin(df_after['_key'])]
    deleted = deleted_keys[cols_to_use].copy()

    # Find added rows (in after but not in before)
    added_keys = df_after[~df_after['_key'].isin(df_before['_key'])]
    added = added_keys[cols_to_use].copy()

    # Find modified rows (same key but different values in compare columns)
    common_keys = set(df_before['_key']) & set(df_after['_key'])

    modified_before = []
    modified_after = []

    for key in common_keys:
        row_before = df_before[df_before['_key'] == key][cols_to_use].iloc[0]
        row_after = df_after[df_after['_key'] == key][cols_to_use].iloc[0]

        # Check if any of the compare columns have changed
        has_changes = False
        for col in compare_columns:
            if str(row_before[col]) != str(row_after[col]):
                has_changes = True
                break

        if has_changes:
            modified_before.append(row_before)
            modified_after.append(row_after)

    modified_before_df = pd.DataFrame(modified_before) if modified_before else pd.DataFrame(columns=cols_to_use)
    modified_after_df = pd.DataFrame(modified_after) if modified_after else pd.DataFrame(columns=cols_to_use)

    return deleted, added, modified_before_df, modified_after_df


def create_excel_output(deleted, added, modified_before, modified_after):
    """Create Excel file with formatted sheets"""
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Write deleted rows
        if not deleted.empty:
            deleted.to_excel(writer, sheet_name='Deleted', index=False)
        else:
            pd.DataFrame(['No deleted rows']).to_excel(writer, sheet_name='Deleted', index=False, header=False)

        # Write added rows
        if not added.empty:
            added.to_excel(writer, sheet_name='Added', index=False)
        else:
            pd.DataFrame(['No added rows']).to_excel(writer, sheet_name='Added', index=False, header=False)

        # Write modified rows (before and after side by side)
        if not modified_before.empty:
            # Rename columns to distinguish before/after
            modified_before_renamed = modified_before.copy()
            modified_after_renamed = modified_after.copy()

            modified_before_renamed.columns = [f"{col} (Before)" for col in modified_before.columns]
            modified_after_renamed.columns = [f"{col} (After)" for col in modified_after.columns]

            modified_combined = pd.concat([modified_before_renamed, modified_after_renamed], axis=1)
            modified_combined.to_excel(writer, sheet_name='Modified', index=False)
        else:
            pd.DataFrame(['No modified rows']).to_excel(writer, sheet_name='Modified', index=False, header=False)

    output.seek(0)
    return output


# Read and store files
if file_before and file_after:
    try:
        # Read files
        df_before = read_file(file_before)
        df_after = read_file(file_after)

        if df_before is None or df_after is None:
            st.error("Error reading files. Please check file formats.")
        else:
            # Store in session state
            st.session_state.df_before = df_before
            st.session_state.df_after = df_after

            # Display file info
            st.success("✅ Files loaded successfully!")
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Before file rows", len(df_before))
                st.metric("Before file columns", len(df_before.columns))
            with col2:
                st.metric("After file rows", len(df_after))
                st.metric("After file columns", len(df_after.columns))

            # Check and display column layout
            st.subheader("📋 Column Layout Verification")

            before_cols = set(df_before.columns)
            after_cols = set(df_after.columns)

            col1, col2 = st.columns(2)

            with col1:
                st.write("**Before file columns:**")
                st.write(list(df_before.columns))

            with col2:
                st.write("**After file columns:**")
                st.write(list(df_after.columns))

            # Check if layouts match
            if before_cols == after_cols:
                st.success("✅ **Layout verification passed!** Both files have the same columns.")

                # Layout confirmation checkbox
                layout_confirmed = st.checkbox(
                    "✓ I confirm that both files share the same layout and are ready for comparison",
                    value=st.session_state.layout_confirmed
                )

                if layout_confirmed:
                    st.session_state.layout_confirmed = True
                    st.markdown("---")

                    # Column selection section
                    st.subheader("🔑 Step 1: Select Identifier Columns")
                    st.markdown(
                        "Choose one or more columns that **uniquely identify each row** (e.g., ID, Employee Number, Product Code, etc.)")

                    key_columns = st.multiselect(
                        "Identifier columns (can select multiple)",
                        options=df_before.columns.tolist(),
                        default=[],
                        help="These columns will be used to match rows between the two files"
                    )

                    if key_columns:
                        st.success(f"✅ Selected {len(key_columns)} identifier column(s): {', '.join(key_columns)}")

                        st.markdown("---")
                        st.subheader("📊 Step 2: Select Columns to Compare")
                        st.markdown(
                            "Choose which columns you want to compare for changes (excludes noise/irrelevant data)")

                        # Get remaining columns (excluding key columns)
                        available_compare_cols = [col for col in df_before.columns if col not in key_columns]

                        compare_columns = st.multiselect(
                            "Columns to compare for changes",
                            options=available_compare_cols,
                            default=available_compare_cols,
                            help="Only changes in these columns will be tracked. Identifier columns are automatically included."
                        )

                        if compare_columns:
                            st.success(f"✅ Will compare {len(compare_columns)} column(s): {', '.join(compare_columns)}")

                            st.markdown("---")

                            # Compare button
                            if st.button("🔍 Compare Files", type="primary", use_container_width=True):
                                with st.spinner("Comparing files..."):
                                    deleted, added, modified_before, modified_after = compare_dataframes(
                                        df_before.copy(), df_after.copy(), key_columns, compare_columns
                                    )

                                    # Display summary
                                    st.subheader("📊 Comparison Summary")
                                    col1, col2, col3 = st.columns(3)

                                    with col1:
                                        st.metric("🗑️ Deleted Rows", len(deleted))
                                    with col2:
                                        st.metric("➕ Added Rows", len(added))
                                    with col3:
                                        st.metric("✏️ Modified Rows", len(modified_before))

                                    # Show previews
                                    if not deleted.empty:
                                        with st.expander("🗑️ View Deleted Rows", expanded=False):
                                            st.dataframe(deleted, use_container_width=True)

                                    if not added.empty:
                                        with st.expander("➕ View Added Rows", expanded=False):
                                            st.dataframe(added, use_container_width=True)

                                    if not modified_before.empty:
                                        with st.expander("✏️ View Modified Rows", expanded=False):
                                            col_a, col_b = st.columns(2)
                                            with col_a:
                                                st.write("**Before:**")
                                                st.dataframe(modified_before, use_container_width=True)
                                            with col_b:
                                                st.write("**After:**")
                                                st.dataframe(modified_after, use_container_width=True)

                                    if deleted.empty and added.empty and modified_before.empty:
                                        st.info("ℹ️ No differences found between the files based on selected columns.")

                                    # Generate Excel output
                                    excel_output = create_excel_output(deleted, added, modified_before, modified_after)

                                    st.markdown("---")
                                    st.download_button(
                                        label="📥 Download Comparison Report (Excel)",
                                        data=excel_output,
                                        file_name="comparison_report.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        type="primary",
                                        use_container_width=True
                                    )

                                    st.success("✅ Comparison complete!")
                        else:
                            st.warning("⚠️ Please select at least one column to compare.")
                    else:
                        st.info("👆 Please select identifier column(s) to proceed.")

            else:
                # Show differences
                st.error("❌ **Layout verification failed!** The files have different columns.")

                only_in_before = before_cols - after_cols
                only_in_after = after_cols - before_cols

                if only_in_before:
                    st.warning(f"**Columns only in 'Before' file:** {', '.join(only_in_before)}")
                if only_in_after:
                    st.warning(f"**Columns only in 'After' file:** {', '.join(only_in_after)}")

                st.info("💡 Please ensure both files have the same column structure before proceeding.")

    except Exception as e:
        st.error(f"❌ An error occurred: {str(e)}")
        st.exception(e)
else:
    st.info("👆 Please upload both files to start the comparison.")
    st.session_state.layout_confirmed = False

# Footer
st.markdown("---")
st.markdown("**How to use:**")
st.markdown("""
1. **Upload files**: Upload your 'before' (baseline) and 'after' (updated) files
2. **Verify layout**: Check that both files have matching column structures
3. **Confirm layout**: Check the confirmation box to proceed
4. **Select identifiers**: Choose one or more columns that uniquely identify each row
5. **Select compare columns**: Choose which columns to compare (exclude noise/irrelevant columns)
6. **Compare**: Click 'Compare Files' to analyze differences
7. **Download**: Get an Excel report with three sheets (Deleted, Added, Modified rows)
""")