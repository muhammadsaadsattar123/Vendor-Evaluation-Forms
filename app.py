import streamlit as st
import openpyxl
import pandas as pd
import io
from datetime import datetime

# Page Configuration
st.set_page_config(
    page_title="Supplier Evaluation Portal",
    page_icon="📋",
    layout="wide"
)

st.title("📋 Supplier Evaluation Form Auto-Processor")
st.write("Master Destination Sheet aur Supplier Forms upload karein. System automatically sara data parse karke master sheet mein append kar dega.")

st.divider()

# File Uploaders
col1, col2 = st.columns(2)

with col1:
    st.subheader("1️⃣ Master Destination Sheet")
    master_file = st.file_uploader("Main Master Excel Sheet Upload Karein", type=["xlsx"], key="master")

with col2:
    st.subheader("2️⃣ Supplier Evaluation Forms")
    forms_files = st.file_uploader("Naye Supplier Forms Upload Karein (Ek ya zyaada)", type=["xlsx"], accept_multiple_files=True, key="forms")

st.divider()

def parse_form_data(ws_form, next_id):
    def get_val(r, c):
        v = ws_form.cell(row=r, column=c).value
        return str(v).strip() if v is not None else ""

    def parse_rating(row, col_num, col_text, alt_col_num=None, alt_col_text=None):
        val_num = get_val(row, col_num)
        val_text = get_val(row, col_text)
        if val_num in ['1', '2', '3']:
            return f"{val_num}-{val_text}" if val_text else val_num
        if alt_col_num:
            val_num_alt = get_val(row, alt_col_num)
            val_text_alt = get_val(row, alt_col_text)
            if val_num_alt in ['1', '2', '3']:
                return f"{val_num_alt}-{val_text_alt}" if val_text_alt else val_num_alt
        return ""

    eval_date = get_val(12, 6).split(' ')[0] if get_val(12, 6) else datetime.now().strftime('%Y-%m-%d')
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    return {
        'Id': next_id,
        'Start time': now_str,
        'Completion time': now_str,
        'Email': f"{get_val(6, 14).lower().replace(' ', '.')}@{get_val(6, 6).lower().replace(' ', '')}.com",
        'Name': get_val(6, 14),
        'Company Name:': get_val(6, 6),
        'Location / Address:': get_val(8, 6),
        'Supplier ID:': get_val(10, 6),
        'Evaluation Date:': eval_date,
        'Contact Person': get_val(6, 14),
        'Contact Number:': get_val(8, 14),
        'Product/Service Provided:': get_val(10, 14),
        'Period Evaluated:1': get_val(12, 14),
        'Evaluation Cycle:': 'Annually',
        'In Business for:': 'More than 5 Years',
        'Registered vendor at TCF for:': 'More than 5 Years',
        'Quality Standard Compliance:': 'Others',

        # Ratings, Measures & Remarks
        'Quality of Product / Service (Weightage%)': get_val(29, 8),
        'Quality of Product / Service: High = 3 (>95%),\xa0Medium = 2 (>80%),\xa0Low = 1 (<80%)': parse_rating(29, 13, 14, 10, 11),
        'Quality of Product / Service:\xa0Measure (If applicable)': get_val(31, 8),
        'Quality of Product / Service (Remarks)': get_val(33, 8),

        'Consistency of Quality (Weightage%)': get_val(36, 8),
        'Consistency of Quality: High = 3 (>95%),\xa0Medium = 2 (>80%),\xa0Low = 1 (<80%)': parse_rating(36, 13, 14, 10, 11),
        'Consistency of Quality:\xa0Measure (If applicable)': get_val(38, 8),
        'Consistency of Quality (Remarks)': get_val(39, 8) or get_val(40, 8),

        'Meets Specifications (Weightage%)': get_val(43, 8),
        'Meets Specifications: High = 3 (>95%),\xa0Medium = 2 (>80%),\xa0Low = 1 (<80%)': parse_rating(43, 10, 11, 13, 14),
        'Meets Specifications:\xa0Measure (If applicable)': get_val(45, 8),
        'Meets Specifications (Remarks)': get_val(47, 8),

        'Pricing compared to market rates (Weightage%)': get_val(55, 8),
        'Pricing compared to market rates: High = 3 (>95%),\xa0Medium = 2 (>80%),\xa0Low = 1 (<80%)': parse_rating(55, 10, 11, 13, 14),
        'Pricing compared to market rates:\xa0Measure (If applicable)': get_val(57, 8),
        'Pricing compared to market rates (Remarks)': get_val(59, 8),

        'Value for Money (Weightage%)': get_val(62, 8),
        'Value for Money: High = 3 (>95%),\xa0Medium = 2 (>80%),\xa0Low = 1 (<80%)': parse_rating(62, 13, 14, 10, 11),
        'Value for Money:\xa0Measure (If applicable)': get_val(64, 8),
        'Value for Money (Remarks)': get_val(66, 8),

        'Discounts/Negotiations Provided (Weightage%)': get_val(69, 8),
        'Discounts/Negotiations Provided: High = 3 (>95%),\xa0Medium = 2 (>80%),\xa0Low = 1 (<80%)': parse_rating(69, 13, 14, 10, 11),
        'Discounts/Negotiations Provided:\xa0Measure (If applicable)': get_val(71, 11) or get_val(71, 8),
        'Discounts/Negotiations Provided (Remarks)': get_val(73, 8),

        'Timely Delivery (Weightage%)': get_val(81, 8),
        'Timely Delivery: High = 3 (>95%),\xa0Medium = 2 (>80%),\xa0Low = 1 (<80%)': parse_rating(81, 13, 14, 10, 11),
        'Timely Delivery:\xa0Measure (If applicable)': get_val(83, 8),
        'Timely Delivery (Remarks)': get_val(85, 8),

        'Condition of Product upon delivery (Weightage%)': get_val(88, 8),
        'Condition of Product upon delivery: High = 3 (>95%),\xa0Medium = 2 (>80%),\xa0Low = 1 (<80%)': parse_rating(88, 13, 14, 10, 11),
        'Condition of Product upon delivery:\xa0Measure (If applicable)': get_val(90, 10) or get_val(90, 8),
        'Condition of Product upon delivery (Remarks)': get_val(91, 8) or get_val(92, 8),

        'Flexibility in Urgent requirements (Weightage%)': get_val(95, 8),
        'Flexibility in Urgent requirements: High = 3 (>95%),\xa0Medium = 2 (>80%),\xa0Low = 1 (<80%)': parse_rating(95, 10, 11, 13, 14),
        'Flexibility in Urgent requirements:\xa0Measure (If applicable)': get_val(97, 10) or get_val(97, 8),
        'Flexibility in Urgent requirements: (Remarks)': get_val(99, 8),

        'Responsiveness to Queries/Concerns (Weightage%)': get_val(108, 8),
        'Responsiveness to Queries/Concerns: High = 3 (>95%),\xa0Medium = 2 (>80%),\xa0Low = 1 (<80%)': parse_rating(108, 10, 11, 13, 14),
        'Responsiveness to Queries/Concerns:\xa0Measure (If applicable)': get_val(110, 8),
        'Responsiveness to Queries/Concerns:\xa0(Remarks)': get_val(112, 8),

        'After-Sales Support/Services (Weightage%)': get_val(115, 8),
        'After-Sales Support/Services: High = 3 (>95%),\xa0Medium = 2 (>80%),\xa0Low = 1 (<80%)': parse_rating(115, 13, 14, 10, 11),
        'After-Sales Support/Services:\xa0Measure (If applicable)': get_val(117, 10) or get_val(117, 8),
        'After-Sales Support/Services: (Remarks)': get_val(119, 8),

        'Professionalism (Weightage%)': get_val(122, 8),
        'Professionalism: High = 3 (>95%),\xa0Medium = 2 (>80%),\xa0Low = 1 (<80%)': parse_rating(122, 10, 11, 13, 14),
        'Professionalism:\xa0Measure (If applicable)': get_val(124, 8),
        'Professionalism: (Remarks)': get_val(126, 8),

        'Adhere to Legal & Ethical standards (Weightage%)': get_val(134, 8),
        'Adhere to Legal & Ethical standards: High = 3 (>95%),\xa0Medium = 2 (>80%),\xa0Low = 1 (<80%)': parse_rating(134, 13, 14, 10, 11),
        'Adhere to Legal & Ethical standards:\xa0Measure (If applicable)': get_val(136, 10) or get_val(136, 8),
        'Adhere to Legal & Ethical standards:\xa0(Remarks)': get_val(137, 8) or get_val(138, 8),

        'Accuracy of invoices & Documents (Weightage%)': get_val(140, 8),
        'Accuracy of invoices & Documents: High = 3 (>95%),\xa0Medium = 2 (>80%),\xa0Low = 1 (<80%)': parse_rating(140, 13, 14, 10, 11),
        'Accuracy of invoices & Documents: Measure (If applicable)': get_val(142, 8),
        'Accuracy of invoices & Documents: (Remarks)': get_val(144, 8),

        'Overall Comments:': get_val(170, 5)
    }

# Process Action
if st.button("🚀 Merge & Update Master Sheet", type="primary", use_container_width=True):
    if not master_file:
        st.error("⚠️ Pehle Master Destination Sheet upload karein!")
    elif not forms_files:
        st.error("⚠️ Kam se kam ek Evaluation Form upload karein!")
    else:
        with st.spinner("Processing forms and merging data..."):
            dest_wb = openpyxl.load_workbook(io.BytesIO(master_file.getvalue()))
            dest_ws = dest_wb.active
            headers = [dest_ws.cell(row=1, column=c).value for c in range(1, dest_ws.max_column + 1)]

            count = 0
            for form_file in forms_files:
                form_wb = openpyxl.load_workbook(io.BytesIO(form_file.getvalue()), data_only=True)
                ws_form = form_wb['LightupEnterprises'] if 'LightupEnterprises' in form_wb.sheetnames else form_wb.active
                
                next_id = dest_ws.max_row
                row_dict = parse_form_data(ws_form, next_id)
                new_row = [row_dict.get(h, "") for h in headers]
                dest_ws.append(new_row)
                count += 1

            output_stream = io.BytesIO()
            dest_wb.save(output_stream)
            output_stream.seek(0)

            st.success(f"🎉 Mubarak ho! {count} supplier form(s) master sheet mein merge ho chuke hain.")

            st.download_button(
                label="📥 Updated Master Sheet Download Karein",
                data=output_stream,
                file_name="Master_Destination_Sheet_Updated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )