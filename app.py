import streamlit as st
import pandas as pd
import os
import glob
import re
from datetime import datetime
from pathlib import Path
import io


class NameSearcher:
    """Name Search Tool reading from fixed folder"""

    def __init__(self, excel_folder="nadiad_chaklasi_excel_database"):
        self.excel_folder = excel_folder

    def extract_part_number(self, file_path):
        """Extract part number from row 6 of the Excel file"""
        try:
            df = pd.read_excel(file_path, sheet_name=0, nrows=10, header=None, dtype=str)

            if len(df) > 5:
                row_6_data = df.iloc[5]

                for cell in row_6_data:
                    if pd.notna(cell):
                        cell_str = str(cell)
                        if ':' in cell_str:
                            part_number = cell_str.split(':')[-1].strip()
                            if part_number:
                                return part_number

            return "N/A"
        except Exception as e:
            return "Error"

    def extract_vidhansabha(self, file_path):
        """Extract Vidhansabha from row 7 of the Excel file"""
        try:
            df = pd.read_excel(file_path, sheet_name=0, nrows=10, header=None, dtype=str)

            if len(df) > 6:
                row_7_data = df.iloc[6]

                for cell in row_7_data:
                    if pd.notna(cell):
                        cell_str = str(cell)
                        if ':' in cell_str:
                            vidhansabha = cell_str.split(':')[-1].strip()
                            if vidhansabha:
                                return vidhansabha

            return "N/A"
        except Exception as e:
            return "Error"

    def extract_row_number(self, matched_content):
        """Extract the first number (before first space) from matched content"""
        if pd.isna(matched_content):
            return ""

        parts = str(matched_content).strip().split()
        if parts:
            return parts[0]
        return ""

    def search_single_excel_file(self, file_path, search_terms, search_names_map):
        """Search for patterns in a single Excel file"""
        results = []

        try:
            part_number = self.extract_part_number(file_path)
            vidhansabha = self.extract_vidhansabha(file_path)
            excel_data = pd.read_excel(file_path, sheet_name=None, dtype=str)

            for sheet_name, df in excel_data.items():
                for row_idx, row in df.iterrows():
                    for col_idx, cell_value in enumerate(row):
                        if pd.notna(cell_value):
                            cell_str = str(cell_value)

                            for pattern in search_terms:
                                if re.search(pattern, cell_str):
                                    row_number = self.extract_row_number(cell_str)
                                    search_name = search_names_map.get(pattern, pattern)

                                    results.append({
                                        'Searched_Name': search_name,
                                        'Vidhansabha': vidhansabha,
                                        'Part_Number': part_number,
                                        'Row_Number': row_number,
                                        'Matched_Content': cell_str
                                    })
                                    break
        except Exception as e:
            st.error(f"Error reading {os.path.basename(file_path)}: {e}")

        return results

    def search_all_excel_files(self, search_terms, search_names_map, all_search_names):
        """Search all Excel files in the fixed folder"""
        excel_files = glob.glob(os.path.join(self.excel_folder, "*.xlsx")) + glob.glob(os.path.join(self.excel_folder, "*.xls"))

        if not excel_files:
            return None, f"No Excel files found in '{self.excel_folder}' folder"

        all_results = []
        found_names = set()
        progress_placeholder = st.empty()

        for idx, file_path in enumerate(excel_files):
            filename = os.path.basename(file_path)
            progress_placeholder.text(f"📄 Searching: {filename}... ({idx + 1}/{len(excel_files)})")

            file_results = self.search_single_excel_file(file_path, search_terms, search_names_map)
            all_results.extend(file_results)

            for result in file_results:
                found_names.add(result['Searched_Name'])

        progress_placeholder.empty()

        # Add "Not Found" entries for names that weren't found
        not_found_names = set(all_search_names) - found_names
        for name in not_found_names:
            all_results.append({
                'Searched_Name': name,
                'Vidhansabha': 'Not Found',
                'Part_Number': 'Not Found',
                'Row_Number': '',
                'Matched_Content': ''
            })

        return all_results, len(excel_files)

    def auto_adjust_column_width(self, worksheet, dataframe):
        """Auto-adjust column widths based on content"""
        for idx, col in enumerate(dataframe.columns):
            max_length = len(str(col))

            # Check content length
            for value in dataframe[col].astype(str):
                max_length = max(max_length, len(value))

            # Set width with some padding (max 80 chars for readability)
            adjusted_width = min(max_length + 2, 80)
            worksheet.column_dimensions[chr(65 + idx)].width = adjusted_width

    def sort_results_by_input_order(self, results_df, input_names_list):
        """Sort results to match the input order of names"""
        # Create a mapping of name to its input order
        name_order_map = {name: idx for idx, name in enumerate(input_names_list)}

        # Add a temporary column for sorting
        results_df['_sort_order'] = results_df['Searched_Name'].map(name_order_map)

        # Sort by this column, then by other criteria for stable sorting
        results_df = results_df.sort_values('_sort_order', kind='stable')

        # Remove the temporary sorting column
        results_df = results_df.drop('_sort_order', axis=1)

        return results_df

    def create_results_excel(self, results, search_terms_display):
        """Create Excel file with results - SAME FORMAT AS offline_app.py"""
        output = io.BytesIO()

        results_df = pd.DataFrame(results)
        # Sort by input order of names
        results_df = self.sort_results_by_input_order(results_df, search_terms_display)

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Sheet 1: Search_Results (all results grouped by name in input order)
            results_df.to_excel(writer, sheet_name='Search_Results', index=False)
            worksheet = writer.sheets['Search_Results']
            self.auto_adjust_column_width(worksheet, results_df)

            # Only create summaries if there are found results
            if len(results) > 0:
                found_results = results_df[results_df['Part_Number'] != 'Not Found']

                if len(found_results) > 0:
                    # Sheet 2: Summary by Name (maintaining input order)
                    name_summary = found_results.groupby('Searched_Name', sort=False).size().reset_index(name='Match_Count')
                    # Reorder based on input
                    name_order_map = {name: idx for idx, name in enumerate(search_terms_display)}
                    name_summary['_sort_order'] = name_summary['Searched_Name'].map(name_order_map)
                    name_summary = name_summary.sort_values('_sort_order', kind='stable')
                    name_summary = name_summary.drop('_sort_order', axis=1)
                    name_summary.to_excel(writer, sheet_name='Summary_by_Name', index=False)
                    worksheet_summary = writer.sheets['Summary_by_Name']
                    self.auto_adjust_column_width(worksheet_summary, name_summary)

                    # Sheet 3: Summary by Part
                    part_summary = found_results.groupby('Part_Number').size().reset_index(name='Match_Count')
                    part_summary = part_summary.sort_values('Match_Count', ascending=False)
                    part_summary.to_excel(writer, sheet_name='Summary_by_Part', index=False)
                    worksheet_part = writer.sheets['Summary_by_Part']
                    self.auto_adjust_column_width(worksheet_part, part_summary)

            # Sheet 4: Search_Terms
            patterns_df = pd.DataFrame({'Search_Terms_Used': search_terms_display})
            patterns_df.to_excel(writer, sheet_name='Search_Terms', index=False)
            worksheet_terms = writer.sheets['Search_Terms']
            self.auto_adjust_column_width(worksheet_terms, patterns_df)

        output.seek(0)
        return output, results_df


def prepare_search_terms(names):
    """Prepare regex search terms from names and create mapping"""
    search_terms = []
    search_names_map = {}

    for name in names:
        pattern1 = f".*{name}.*"
        pattern2 = f"(?i).*{name}.*"
        search_terms.append(pattern1)
        search_terms.append(pattern2)
        search_names_map[pattern1] = name
        search_names_map[pattern2] = name

    return search_terms, search_names_map


def main():
    st.set_page_config(
        page_title="Name Search Tool",
        page_icon="🔍",
        layout="wide"
    )

    st.title("🔍 Name Search Tool")
    st.markdown("Search for names in Excel files easily!")

    EXCEL_FOLDER = "nadiad_chaklasi_excel_database"

    # Check if folder exists
    if not os.path.exists(EXCEL_FOLDER):
        st.error(f"❌ Excel folder '{EXCEL_FOLDER}' not found!")
        st.info("Please add the 'excel_output' folder with Excel files.")
        st.stop()

    # Count files
    excel_files = glob.glob(os.path.join(EXCEL_FOLDER, "*.xlsx")) + glob.glob(os.path.join(EXCEL_FOLDER, "*.xls"))

    if not excel_files:
        st.error(f"❌ No Excel files found in '{EXCEL_FOLDER}' folder!")
        st.info("Please add Excel files to the 'excel_output' folder.")
        st.stop()

    st.success(f"✅ {len(excel_files)} Excel files loaded and ready!")

    with st.expander(f"📂 View Available Files ({len(excel_files)} files)"):
        for i, file in enumerate(excel_files, 1):
            st.write(f"{i}. {os.path.basename(file)}")

    st.markdown("---")

    searcher = NameSearcher(EXCEL_FOLDER)

    if 'search_terms' not in st.session_state:
        st.session_state.search_terms = None
    if 'search_terms_display' not in st.session_state:
        st.session_state.search_terms_display = []
    if 'search_names_map' not in st.session_state:
        st.session_state.search_names_map = {}
    if 'results_data' not in st.session_state:
        st.session_state.results_data = None
    if 'input_filename' not in st.session_state:
        st.session_state.input_filename = None

    # Sidebar
    st.sidebar.header("📋 Provide Names to Search")

    input_method = st.sidebar.radio(
        "Select input method:",
        ["Type Names Manually", "Upload Text File (.txt)", "Upload Excel File"],
    )

    if input_method == "Type Names Manually":
        st.sidebar.markdown("**Enter names (one per line):**")
        manual_input = st.sidebar.text_area(
            "Type here:",
            height=250,
            placeholder="પટેલ\nશાહ\nPatel\nShah",
        )

        if st.sidebar.button("✅ Load Names", type="primary", use_container_width=True):
            if manual_input.strip():
                lines = [line.strip() for line in manual_input.splitlines() if line.strip()]
                search_terms, search_names_map = prepare_search_terms(lines)
                st.session_state.search_terms = search_terms
                st.session_state.search_terms_display = lines
                st.session_state.search_names_map = search_names_map
                st.session_state.results_data = None
                st.session_state.input_filename = "manual_input"
                st.sidebar.success(f"✅ {len(lines)} names loaded!")
                st.rerun()
            else:
                st.sidebar.error("⚠️ Enter at least one name")

    elif input_method == "Upload Text File (.txt)":
        st.sidebar.markdown("**Upload text file:**")
        txt_file = st.sidebar.file_uploader("Choose file", type=['txt'], key="txt")

        if txt_file and st.sidebar.button("✅ Load", type="primary", use_container_width=True):
            try:
                lines = txt_file.read().decode('utf-8').splitlines()
                lines = [line.strip() for line in lines if line.strip()]
                search_terms, search_names_map = prepare_search_terms(lines)
                st.session_state.search_terms = search_terms
                st.session_state.search_terms_display = lines
                st.session_state.search_names_map = search_names_map
                st.session_state.results_data = None
                st.session_state.input_filename = Path(txt_file.name).stem
                st.sidebar.success(f"✅ {len(lines)} names loaded!")
                st.rerun()
            except Exception as e:
                st.sidebar.error(f"Error: {e}")

    elif input_method == "Upload Excel File":
        st.sidebar.markdown("**Upload Excel file:**")
        excel_input = st.sidebar.file_uploader("Choose file", type=['xlsx', 'xls'], key="excel")

        if excel_input:
            try:
                # Read Excel WITHOUT header
                df = pd.read_excel(excel_input, dtype=str, header=None)
                num_columns = len(df.columns)

                # Single column - auto select
                if num_columns == 1:
                    st.sidebar.info(f"✅ Using the only column")
                    col_idx = 0

                    if st.sidebar.button("✅ Load", type="primary", use_container_width=True):
                        values = df[col_idx].dropna().unique()
                        display_names = [str(v).strip() for v in values if str(v).strip()]
                        search_terms, search_names_map = prepare_search_terms(display_names)
                        st.session_state.search_terms = search_terms
                        st.session_state.search_terms_display = display_names
                        st.session_state.search_names_map = search_names_map
                        st.session_state.results_data = None
                        st.session_state.input_filename = Path(excel_input.name).stem
                        st.sidebar.success(f"✅ {len(values)} names loaded!")
                        st.rerun()

                else:
                    # Multiple columns - numbered list
                    st.sidebar.markdown("**📋 Available Columns:**")
                    for idx in range(num_columns):
                        col_sample = df[idx].dropna().iloc[0] if len(df[idx].dropna()) > 0 else f"Column {idx+1}"
                        st.sidebar.write(f"**{idx+1}.** {col_sample} (Sample)")

                    col_choice = st.sidebar.text_input("Enter column number to use:", placeholder="1")

                    if st.sidebar.button("✅ Load", type="primary", use_container_width=True):
                        try:
                            col_idx = int(col_choice) - 1
                            if 0 <= col_idx < num_columns:
                                values = df[col_idx].dropna().unique()
                                display_names = [str(v).strip() for v in values if str(v).strip()]
                                search_terms, search_names_map = prepare_search_terms(display_names)
                                st.session_state.search_terms = search_terms
                                st.session_state.search_terms_display = display_names
                                st.session_state.search_names_map = search_names_map
                                st.session_state.results_data = None
                                st.session_state.input_filename = Path(excel_input.name).stem
                                st.sidebar.success(f"✅ {len(values)} names loaded!")
                                st.rerun()
                            else:
                                st.sidebar.error(f"❌ Invalid! Enter number between 1 and {num_columns}")
                        except ValueError:
                            st.sidebar.error("❌ Please enter a valid number!")

            except Exception as e:
                st.sidebar.error(f"Error: {e}")

    if st.session_state.search_terms:
        st.sidebar.markdown("---")
        st.sidebar.success(f"✅ **{len(st.session_state.search_terms_display)} names ready!**")

        with st.sidebar.expander("👁️ View"):
            for i, name in enumerate(st.session_state.search_terms_display, 1):
                st.write(f"{i}. {name}")

        if st.sidebar.button("🗑️ Clear", use_container_width=True):
            st.session_state.search_terms = None
            st.session_state.search_terms_display = []
            st.session_state.search_names_map = {}
            st.session_state.results_data = None
            st.session_state.input_filename = None
            st.rerun()

    # Main area
    if st.session_state.search_terms:
        st.header("🔍 Ready to Search!")
        st.info(f"Will search in {len(excel_files)} Excel files")

        if st.button("🚀 START SEARCH", type="primary", use_container_width=True):
            with st.spinner("Searching..."):
                results, file_count = searcher.search_all_excel_files(
                    st.session_state.search_terms,
                    st.session_state.search_names_map,
                    st.session_state.search_terms_display
                )

                if results is None:
                    st.error(file_count)
                else:
                    st.session_state.results_data = {
                        'results': results,
                        'file_count': file_count
                    }
                    st.rerun()

    # Display results if available
    if st.session_state.results_data:
        results = st.session_state.results_data['results']
        file_count = st.session_state.results_data['file_count']

        st.markdown("---")
        st.header("📊 Results")

        if results:
            results_df = pd.DataFrame(results)
            # Sort by input order of names
            results_df = searcher.sort_results_by_input_order(results_df, st.session_state.search_terms_display)

            found_count = len(results_df[results_df['Part_Number'] != 'Not Found'])
            not_found_count = len(results_df[results_df['Part_Number'] == 'Not Found'])

            st.success(f"🎉 Found {found_count} matches!")
            if not_found_count > 0:
                st.warning(f"⚠️ {not_found_count} names not found")

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Matches", found_count)
            with col2:
                st.metric("Not Found", not_found_count)
            with col3:
                st.metric("Files Searched", file_count)
            with col4:
                unique_parts = results_df[results_df['Part_Number'] != 'Not Found']['Part_Number'].nunique()
                st.metric("Unique Parts", unique_parts)

            # DOWNLOAD BUTTON AT TOP
            st.markdown("### 📥 Download Results")
            excel_output, _ = searcher.create_results_excel(results, st.session_state.search_terms_display)
            output_filename = f"{st.session_state.input_filename}_output.xlsx"

            st.download_button(
                label="📥 Download Results (Excel with Auto-Adjusted Columns)",
                data=excel_output,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )

            st.markdown("---")

            # Results table
            st.subheader("📋 Search Results (Grouped by Name)")
            st.dataframe(
                results_df,
                use_container_width=True,
                height=400,
                column_config={
                    "Searched_Name": "🔍 Searched Name",
                    "Vidhansabha": "🏛️ Vidhansabha",
                    "Part_Number": "🔢 Part Number",
                    "Row_Number": "📍 Row Number",
                    "Matched_Content": "✅ Matched Content",
                },
                hide_index=True
            )

            # Summaries
            col1, col2 = st.columns(2)

            with col1:
                st.subheader("📈 Matches by Name")
                found_results = results_df[results_df['Part_Number'] != 'Not Found']
                if len(found_results) > 0:
                    name_summary = found_results.groupby('Searched_Name', sort=False).size().reset_index(name='Matches')
                    # Reorder based on input
                    name_order_map = {name: idx for idx, name in enumerate(st.session_state.search_terms_display)}
                    name_summary['_sort_order'] = name_summary['Searched_Name'].map(name_order_map)
                    name_summary = name_summary.sort_values('_sort_order', kind='stable')
                    name_summary = name_summary.drop('_sort_order', axis=1)
                    st.dataframe(name_summary, hide_index=True, use_container_width=True)

            with col2:
                st.subheader("📊 Matches by Part")
                if len(found_results) > 0:
                    part_summary = found_results.groupby('Part_Number').size().reset_index(name='Matches')
                    st.dataframe(part_summary.sort_values('Matches', ascending=False), hide_index=True, use_container_width=True)

        else:
            st.warning("❌ No matches found")

    elif not st.session_state.search_terms:
        st.info("👈 **Please provide names to search (sidebar)**")
        st.markdown("### 📝 How to use:")
        st.markdown("1. Choose input method from sidebar")
        st.markdown("2. Load the names to search")
        st.markdown("3. Click START SEARCH")
        st.markdown("4. Download Excel output with auto-adjusted columns")


if __name__ == "__main__":
    main()
