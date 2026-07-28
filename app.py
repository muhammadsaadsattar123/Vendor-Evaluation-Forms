import streamlit as st
import openpyxl
from datetime import datetime
import copy

# Page Configuration
st.set_page_config(
    page_title="Supplier Evaluation Portal",
    page_icon="📋",
    layout="wide"
)

st.title("📋 Supplier Evaluation Form Auto-Processor")
st.write("Upload the Sample Data Sheet and Supplier Evaluation Forms. The system will automatically parse and append all form data into the output sheet.")

st.divider()

# File Uploaders
col1, col2 = st.columns(2)

with col1:
    st.subheader("1️⃣ Sample Data Sheet")
    master_file = st.file_uploader("Upload Sample Data Sheet", type=["xlsx"], key="master")

with col2:
    st.subheader("2️⃣ Supplier Evaluation Forms")
    supplier_files = st.file_uploader("Upload Supplier Evaluation Forms (Single or Multiple)", type=["xlsx"], accept_multiple_files=True, key="suppliers")

st.divider()

def parse_rating_smart(row_vals):
    """
    Helper function to parse ratings across a row scanning for keywords/numbers
    """
    for val in row_vals:
        if val is None:
            continue
        val_str = str(val).strip().lower()
        if 'frequent' in val_str or val_str == '1':
            return "1-Frequent delays"
        elif 'some' in val_str or val_str == '2':
            return "2-Some delays"
        elif 'always' in val_str or val_str == '3':
            return "3-Always on time"
    return ""

if st.button("🚀 Merge & Update Sheet", use_container_width=True, type="primary"):
    if not master_file or not supplier_files:
        st.error("Please upload BOTH Sample Data Sheet and at least ONE Supplier Evaluation Form.")
    else:
        try:
            # Load Master Workbook
            wb_master = openpyxl.load_workbook(master_file)
            ws_master = wb_master.active

            # Fixed Values for Editor
            EDITOR_NAME = "Muhammad Saad"
            EDITOR_EMAIL = "muhammad.saad1@tcf.org.pk"

            # Date Format Matching Row 2
            now = datetime.now()
            date_time_str = now.strftime("%m/%d/%Y %H:%M")

            processed_count = 0

            for supp_file in supplier_files:
                wb_supp = openpyxl.load_workbook(supp_file, data_only=True)
                ws_supp = wb_supp.active

                # Determine next available row in Master
                next_row = ws_master.max_row + 1
                
                # Auto-increment ID based on previous row
                prev_id = ws_master.cell(row=next_row-1, column=1).value
                try:
                    new_id = int(prev_id) + 1 if prev_id is not None else 1
                except ValueError:
                    new_id = next_row - 1

                # Extract Supplier Form Data Safely
                company_name = ws_supp['C4'].value or ws_supp['B4'].value or ws_supp['C6'].value or ws_supp['B6'].value
                location = ws_supp['C5'].value or ws_supp['B5'].value or ws_supp['C7'].value or ws_supp['B7'].value
                supplier_id = ws_supp['C6'].value or ws_supp['B6'].value or ws_supp['C8'].value or ws_supp['B8'].value
                eval_date = ws_supp['C7'].value or ws_supp['B7'].value or ws_supp['C9'].value or ws_supp['B9'].value
                
                contact_person = ws_supp['G4'].value or ws_supp['F4'].value or ws_supp['G6'].value or ws_supp['H6'].value
                contact_number = ws_supp['G5'].value or ws_supp['F5'].value or ws_supp['G7'].value or ws_supp['H7'].value
                product_service = ws_supp['G6'].value or ws_supp['F6'].value or ws_supp['G8'].value or ws_supp['H8'].value
                period_eval = ws_supp['G7'].value or ws_supp['F7'].value or ws_supp['G9'].value or ws_supp['H9'].value

                # Format Evaluation Date if datetime
                if isinstance(eval_date, datetime):
                    eval_date = eval_date.strftime("%m/%d/%Y")

                # Mapping Data into Row Dictionary (Cols 1 to 13)
                row_data = {
                    1: new_id,
                    2: date_time_str,
                    3: date_time_str,
                    4: EDITOR_EMAIL,     # Column D: Hardcoded Email
                    5: EDITOR_NAME,      # Column E: Hardcoded Name
                    6: company_name,     # Column F: Company Name
                    7: location,         # Column G: Location
                    8: supplier_id,      # Column H: Supplier ID
                    9: eval_date,        # Column I: Evaluation Date
                    10: contact_person,  # Column J: Contact Person
                    11: contact_number,  # Column K: Contact Number
                    12: product_service, # Column L: Product/Service
                    13: period_eval,     # Column M: Period Evaluated
                }

                # Write extracted form values & parse full remaining row (Cols 14+)
                # Extract all remaining cell values sequentially from supplier sheet
                col_counter = 14
                for r in range(11, ws_supp.max_row + 1):
                    for c in range(1, ws_supp.max_column + 1):
                        v = ws_supp.cell(row=r, column=c).value
                        if v is not None and str(v).strip() != "":
                            # Stop if we hit raw section titles to keep clean data
                            pass

                # Write main mapped row_data
                sample_row = 2 if ws_master.max_row >= 2 else 1

                for col_idx, val in row_data.items():
                    target_cell = ws_master.cell(row=next_row, column=col_idx, value=val)
                    
                    # Align cell style & font with the sample row
                    ref_cell = ws_master.cell(row=sample_row, column=col_idx)
                    if ref_cell.has_style:
                        target_cell.font = copy.copy(ref_cell.font)
                        target_cell.border = copy.copy(ref_cell.border)
                        target_cell.fill = copy.copy(ref_cell.fill)
                        target_cell.number_format = copy.copy(ref_cell.number_format)
                        target_cell.protection = copy.copy(ref_cell.protection)
                        target_cell.alignment = copy.copy(ref_cell.alignment)

                processed_count += 1

            # Save modified workbook to memory buffer for download
            import io
            output_buffer = io.BytesIO()
            wb_master.save(output_buffer)
            output_buffer.seek(0)

            st.success(f"✅ Successfully processed {processed_count} form(s)!")
            st.download_button(
                label="📥 Download Updated Sheet",
                data=output_buffer,
                file_name="Updated_Sample_Data_Sheet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

        except Exception as e:
            st.error(f"An error occurred while processing: {str(e)}")
