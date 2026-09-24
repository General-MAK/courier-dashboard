import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import zipfile
from io import BytesIO

st.set_page_config(page_title="Courier Expense Reporter", layout="wide")

st.title("📦 Courier Expense Summary & Report Generator")
st.write("Upload your monthly courier detail report to generate executive summaries and download an Excel dashboard.")

uploaded_file = st.file_uploader("Upload Courier Detail Excel File (.xlsx)", type=["xlsx"])

def parse_corrupt_or_standard_excel(file_bytes):
    # Fallback XML parser if stylesheet is malformed
    try:
        df = pd.read_excel(file_bytes, skiprows=2)
        if 'Sales Executive' in df.columns:
            return df
    except Exception:
        pass
    
    file_bytes.seek(0)
    with zipfile.ZipFile(file_bytes, 'r') as z:
        # Read shared strings
        strings = []
        try:
            ss_xml = z.read('xl/sharedStrings.xml')
            root = ET.fromstring(ss_xml)
            ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            for si in root.findall('m:si', ns):
                t = si.find('m:t', ns)
                if t is not None:
                    strings.append(t.text)
                else:
                    parts = [r.find('m:t', ns).text for r in si.findall('m:r', ns) if r.find('m:t', ns) is not None]
                    strings.append("".join(parts))
        except KeyError:
            pass

        sheet_xml = z.read('xl/worksheets/sheet1.xml')
        root = ET.fromstring(sheet_xml)
        ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        sheetData = root.find('m:sheetData', ns)
        rows = []
        for row in sheetData.findall('m:row', ns):
            row_data = []
            for c in row.findall('m:c', ns):
                v = c.find('m:v', ns)
                val = v.text if v is not None else None
                t = c.get('t')
                if t == 's' and val is not None:
                    val = strings[int(val)]
                elif val is not None:
                    try:
                        val = float(val)
                    except ValueError:
                        pass
                row_data.append(val)
            rows.append(row_data)

    df = pd.DataFrame(rows[3:], columns=rows[2])
    return df

if uploaded_file is not None:
    try:
        raw_df = parse_corrupt_or_standard_excel(uploaded_file)
        raw_df = raw_df.dropna(subset=['Sales Executive']).copy()
        raw_df['Budget Utilized'] = pd.to_numeric(raw_df['Budget Utilized'], errors='coerce').fillna(0)
        raw_df['Monthly Budget'] = pd.to_numeric(raw_df['Monthly Budget'], errors='coerce').fillna(0)

        # Month Filter
        available_months = raw_df['Month'].dropna().unique().tolist()
        selected_months = st.multiselect("Select Month(s) for the Summary Report:", options=available_months, default=available_months)

        if not selected_months:
            st.warning("Please select at least one month.")
        else:
            df = raw_df[raw_df['Month'].isin(selected_months)]

            # Computations
            table1_exec = df.groupby('Sales Executive')['Budget Utilized'].sum().reset_index().rename(columns={'Budget Utilized': 'Budget Utilized (PKR)'})
            
            table2_shipments = df.groupby('Sales Executive').size().reset_index(name='Total Shipments')
            
            table3_budget = df.groupby('Sales Executive').agg(
                Budget_Allocated=('Monthly Budget', 'first'),
                Budget_Utilized=('Budget Utilized', 'sum')
            ).reset_index()
            table3_budget['Difference (Under Budget)'] = table3_budget['Budget_Allocated'] - table3_budget['Budget_Utilized']

            table4_region = df.groupby('Country').agg(
                Budget_Utilized=('Budget Utilized', 'sum'),
                No_of_Shipments=('Country', 'count')
            ).reset_index().sort_values('Budget_Utilized', ascending=False)

            table5_customer = df.groupby('Customer').agg(
                Budget_Utilized=('Budget Utilized', 'sum'),
                No_of_Shipments=('Customer', 'count')
            ).reset_index().sort_values('Budget_Utilized', ascending=False)

            # Display on Screen
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("1. Expense Summary by Executive")
                st.dataframe(table1_exec, use_container_width=True)

                st.subheader("2. Shipments by Executive")
                st.dataframe(table2_shipments, use_container_width=True)

            with col2:
                st.subheader("3. Executive Budget Summary")
                st.dataframe(table3_budget, use_container_width=True)

            st.subheader("Executive Budget Comparison")
            st.bar_chart(table3_budget.set_index('Sales Executive')[['Budget_Allocated', 'Budget_Utilized']])

            col3, col4 = st.columns(2)
            with col3:
                st.subheader("4. Spend by Region")
                st.dataframe(table4_region, use_container_width=True)
                st.bar_chart(table4_region.set_index('Country')['Budget_Utilized'].head(10))

            with col4:
                st.subheader("5. Top Customers by Spend")
                st.dataframe(table5_customer, use_container_width=True)
                st.bar_chart(table5_customer.set_index('Customer')['Budget_Utilized'].head(10))

            # Generate Formatted Excel for Download
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                workbook = writer.book
                header_format = workbook.add_format({'bold': True, 'bg_color': '#1F497D', 'font_color': 'white', 'border': 1})
                currency_format = workbook.add_format({'num_format': 'PKR #,##0', 'border': 1})
                num_format = workbook.add_format({'num_format': '#,##0', 'border': 1})
                text_format = workbook.add_format({'border': 1})

                # Write Dashboard Tab
                ws = workbook.add_worksheet('Executive_Dashboard')
                writer.sheets['Executive_Dashboard'] = ws

                def write_table(start_row, start_col, title, df_data, is_currency=True):
                    ws.write(start_row, start_col, title, workbook.add_format({'bold': True, 'font_size': 12}))
                    for c_idx, col_name in enumerate(df_data.columns):
                        ws.write(start_row + 1, start_col + c_idx, col_name, header_format)
                    for r_idx, row in df_data.iterrows():
                        for c_idx, val in enumerate(row):
                            fmt = text_format if isinstance(val, str) else (currency_format if is_currency and c_idx > 0 else num_format)
                            ws.write(start_row + 2 + r_idx, start_col + c_idx, val, fmt)

                write_table(1, 0, "Expense Summary by Executive", table1_exec)
                write_table(len(table1_exec) + 4, 0, "Shipments by Executive", table2_shipments, is_currency=False)
                write_table(1, 4, "Executive Budget Summary", table3_budget)
                write_table(len(table3_budget) + 6, 0, "Courier Spend by Region", table4_region)
                write_table(len(table3_budget) + 6, 4, "Top Customers by Spend", table5_customer)

                ws.set_column('A:A', 25)
                ws.set_column('B:D', 20)
                ws.set_column('E:E', 25)
                ws.set_column('F:H', 20)

            st.download_button(
                label="📥 Download Formatted Excel Report (.xlsx)",
                data=excel_buffer.getvalue(),
                file_name=f"Courier_Summary_{'_'.join(selected_months)}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error reading report: {e}")
