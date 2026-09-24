import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import zipfile
from io import BytesIO

st.set_page_config(page_title="Courier Expense Reporter", layout="wide")

st.title("📦 Courier Expense Summary & Report Generator")
st.write("Upload your monthly courier detail report to generate executive summaries and export a formatted Excel file with native charts.")

uploaded_file = st.file_uploader("Upload Courier Detail Excel File (.xlsx)", type=["xlsx"])

def parse_excel_safely(file_bytes):
    try:
        df = pd.read_excel(file_bytes, skiprows=2)
        if 'Sales Executive' in df.columns:
            return df
    except Exception:
        pass

    file_bytes.seek(0)
    with zipfile.ZipFile(file_bytes, 'r') as z:
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
        raw_df = parse_excel_safely(uploaded_file)
        raw_df = raw_df.dropna(subset=['Sales Executive']).copy()
        raw_df['Budget Utilized'] = pd.to_numeric(raw_df['Budget Utilized'], errors='coerce').fillna(0)
        raw_df['Monthly Budget'] = pd.to_numeric(raw_df['Monthly Budget'], errors='coerce').fillna(0)

        # Month Filter
        available_months = raw_df['Month'].dropna().unique().tolist()
        selected_months = st.multiselect("Select Month(s) for Report:", options=available_months, default=available_months)

        if not selected_months:
            st.warning("Please select at least one month.")
        else:
            df = raw_df[raw_df['Month'].isin(selected_months)]

            # 1. Expense Summary by Executive
            t1_exec = df.groupby('Sales Executive')['Budget Utilized'].sum().reset_index().sort_values('Budget Utilized', ascending=False)

            # 2. Shipments by Executive
            t2_shipments = df.groupby('Sales Executive').size().reset_index(name='Couriers').sort_values('Couriers', ascending=False)

            # 3. Budget Summary
            t3_budget = df.groupby('Sales Executive').agg(
                Budget_Allocated=('Monthly Budget', 'first'),
                Budget_Utilized=('Budget Utilized', 'sum')
            ).reset_index().sort_values('Budget_Allocated', ascending=False)
            t3_budget['Variance'] = t3_budget['Budget_Allocated'] - t3_budget['Budget_Utilized']

            # 4. Region Wise
            t4_region = df.groupby('Country').agg(
                Budget_Utilized=('Budget Utilized', 'sum'),
                No_of_Shipments=('Country', 'count')
            ).reset_index().sort_values('Budget_Utilized', ascending=False)

            # 5. Customer Wise
            t5_customer = df.groupby('Customer').agg(
                Budget_Utilized=('Budget Utilized', 'sum'),
                No_of_Shipments=('Customer', 'count')
            ).reset_index().sort_values('Budget_Utilized', ascending=False)

            # Display on Screen
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("1. Expense Summary by Executive")
                st.dataframe(t1_exec.style.format({'Budget Utilized': 'PKR {:,.2f}'}), use_container_width=True)

                st.subheader("2. Shipments by Executive")
                st.dataframe(t2_shipments, use_container_width=True)

            with col2:
                st.subheader("3. Executive Budget Summary")
                st.dataframe(t3_budget.style.format({'Budget_Allocated': 'PKR {:,.2f}', 'Budget_Utilized': 'PKR {:,.2f}', 'Variance': 'PKR {:,.2f}'}), use_container_width=True)

            st.subheader("Executive Budget Comparison")
            st.bar_chart(t3_budget.set_index('Sales Executive')[['Budget_Allocated', 'Budget_Utilized']])

            col3, col4 = st.columns(2)
            with col3:
                st.subheader("4. Courier Spend by Region")
                st.dataframe(t4_region.style.format({'Budget_Utilized': 'PKR {:,.2f}'}), use_container_width=True)

            with col4:
                st.subheader("5. Top Customers by Spend")
                st.dataframe(t5_customer.head(15).style.format({'Budget_Utilized': 'PKR {:,.2f}'}), use_container_width=True)

            # -------------------------------------------------------------
            # BUILD NATIVE EXCEL WORKBOOK WITH EMBEDDED CHARTS (XLSXWRITER)
            # -------------------------------------------------------------
            excel_buffer = BytesIO()
            workbook = pd.ExcelWriter(excel_buffer, engine='xlsxwriter').book

            # Worksheet 1: Page 1 - Executive Summary
            ws1 = workbook.add_worksheet('Page 1 - Executive Summary')
            ws1.set_paper(9) # A4
            ws1.set_portrait()
            ws1.fit_to_pages(1, 0)

            # Worksheet 2: Page 2 - Region & Customer
            ws2 = workbook.add_worksheet('Page 2 - Region & Customer')
            ws2.set_paper(9) # A4
            ws2.set_portrait()
            ws2.fit_to_pages(1, 0)

            # Styles
            title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#1F497D'})
            sec_hdr_fmt = workbook.add_format({'bold': True, 'font_size': 11, 'bg_color': '#D9E1F2', 'font_color': '#1F497D', 'border': 1})
            th_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F497D', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            cell_txt = workbook.add_format({'border': 1, 'valign': 'vcenter'})
            cell_num = workbook.add_format({'border': 1, 'num_format': '#,##0', 'align': 'right', 'valign': 'vcenter'})
            cell_curr = workbook.add_format({'border': 1, 'num_format': 'PKR #,##0', 'align': 'right', 'valign': 'vcenter'})
            total_txt = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#F2F2F2', 'valign': 'vcenter'})
            total_curr = workbook.add_format({'bold': True, 'border': 1, 'num_format': 'PKR #,##0', 'bg_color': '#F2F2F2', 'align': 'right', 'valign': 'vcenter'})
            total_num = workbook.add_format({'bold': True, 'border': 1, 'num_format': '#,##0', 'bg_color': '#F2F2F2', 'align': 'right', 'valign': 'vcenter'})

            # --- POPULATE PAGE 1 ---
            month_label = ", ".join(selected_months)
            ws1.write(0, 0, f"Courier Budget Detail Report - {month_label}", title_fmt)

            # Table 1: Expense Summary by Executive
            ws1.write(2, 0, "Table 1: Expense Summary by Executive", sec_hdr_fmt)
            ws1.write(3, 0, "Executive", th_fmt)
            ws1.write(3, 1, "Spend (PKR)", th_fmt)
            ws1.write(3, 2, "Grand Total", th_fmt)

            r = 4
            for _, row in t1_exec.iterrows():
                ws1.write(r, 0, row['Sales Executive'], cell_txt)
                ws1.write(r, 1, row['Budget Utilized'], cell_curr)
                ws1.write_formula(r, 2, f"=B{r+1}", cell_curr)
                r += 1
            ws1.write(r, 0, "Grand Total", total_txt)
            ws1.write_formula(r, 1, f"=SUM(B5:B{r})", total_curr)
            ws1.write_formula(r, 2, f"=SUM(C5:C{r})", total_curr)

            # Table 2: Shipments by Executive
            ws1.write(r + 2, 0, "Table 2: Shipments by Executive", sec_hdr_fmt)
            ws1.write(r + 3, 0, "Executive", th_fmt)
            ws1.write(r + 3, 1, "Couriers", th_fmt)

            r2 = r + 4
            start_r2 = r2 + 1
            for _, row in t2_shipments.iterrows():
                ws1.write(r2, 0, row['Sales Executive'], cell_txt)
                ws1.write(r2, 1, row['Couriers'], cell_num)
                r2 += 1
            ws1.write(r2, 0, "Grand Total", total_txt)
            ws1.write_formula(r2, 1, f"=SUM(B{start_r2}:B{r2})", total_num)

            # Table 3: Executive Budget Summary
            ws1.write(2, 4, "Table 3: Executive Budget Summary", sec_hdr_fmt)
            ws1.write(3, 4, "Executive", th_fmt)
            ws1.write(3, 5, "Budget Allocated", th_fmt)
            ws1.write(3, 6, "Budget Utilized", th_fmt)
            ws1.write(3, 7, "Variance", th_fmt)

            r3 = 4
            for _, row in t3_budget.iterrows():
                ws1.write(r3, 4, row['Sales Executive'], cell_txt)
                ws1.write(r3, 5, row['Budget_Allocated'], cell_curr)
                ws1.write(r3, 6, row['Budget_Utilized'], cell_curr)
                ws1.write_formula(r3, 7, f"=F{r3+1}-G{r3+1}", cell_curr)
                r3 += 1
            ws1.write(r3, 4, "Grand Total", total_txt)
            ws1.write_formula(r3, 5, f"=SUM(F5:F{r3})", total_curr)
            ws1.write_formula(r3, 6, f"=SUM(G5:G{r3})", total_curr)
            ws1.write_formula(r3, 7, f"=SUM(H5:H{r3})", total_curr)

            # Chart 1: Executive Budget Comparison (Column Chart)
            chart1 = workbook.add_chart({'type': 'column'})
            chart1.add_series({
                'name': "='Page 1 - Executive Summary'!$F$4",
                'categories': f"='Page 1 - Executive Summary'!$E$5:$E${r3}",
                'values': f"='Page 1 - Executive Summary'!$F$5:$F${r3}",
                'fill': {'color': '#1F497D'},
            })
            chart1.add_series({
                'name': "='Page 1 - Executive Summary'!$G$4",
                'categories': f"='Page 1 - Executive Summary'!$E$5:$E${r3}",
                'values': f"='Page 1 - Executive Summary'!$G$5:$G${r3}",
                'fill': {'color': '#00A0DC'},
            })
            chart1.set_title({'name': 'Executive Budget Comparison: Allocated vs. Utilized'})
            chart1.set_x_axis({'name': 'Executive'})
            chart1.set_y_axis({'name': 'Amount (PKR)', 'major_gridlines': {'visible': True}})
            chart1.set_size({'width': 540, 'height': 320})
            ws1.insert_chart('E13', chart1)

            ws1.set_column('A:A', 22)
            ws1.set_column('B:C', 16)
            ws1.set_column('D:D', 4)
            ws1.set_column('E:E', 22)
            ws1.set_column('F:H', 18)

            # --- POPULATE PAGE 2 ---
            ws2.write(0, 0, f"Courier Spend by Region & Customer - {month_label}", title_fmt)

            # Table 4: Region Wise
            ws2.write(2, 0, "Table 4: Region Wise", sec_hdr_fmt)
            ws2.write(3, 0, "Country", th_fmt)
            ws2.write(3, 1, "Budget Utilized (Cost)", th_fmt)
            ws2.write(3, 2, "No. of Shipments", th_fmt)

            r4 = 4
            for _, row in t4_region.iterrows():
                ws2.write(r4, 0, row['Country'], cell_txt)
                ws2.write(r4, 1, row['Budget_Utilized'], cell_curr)
                ws2.write(r4, 2, row['No_of_Shipments'], cell_num)
                r4 += 1
            ws2.write(r4, 0, "Grand Total", total_txt)
            ws2.write_formula(r4, 1, f"=SUM(B5:B{r4})", total_curr)
            ws2.write_formula(r4, 2, f"=SUM(C5:C{r4})", total_num)

            # Chart 2: Region Bar Chart
            chart2 = workbook.add_chart({'type': 'bar'})
            chart2.add_series({
                'name': 'Budget Utilized',
                'categories': f"='Page 2 - Region & Customer'!$A$5:$A${r4}",
                'values': f"='Page 2 - Region & Customer'!$B$5:$B${r4}",
                'fill': {'color': '#1F497D'},
            })
            chart2.set_title({'name': 'Courier Spend by Region (Top Destinations)'})
            chart2.set_x_axis({'name': 'Budget Utilized (PKR)'})
            chart2.set_size({'width': 480, 'height': 280})
            ws2.insert_chart('E3', chart2)

            # Table 5: Customer Wise
            r5_start = max(r4 + 3, 18)
            ws2.write(r5_start, 0, "Table 5: Customer Wise (Top Customers)", sec_hdr_fmt)
            ws2.write(r5_start + 1, 0, "Customer", th_fmt)
            ws2.write(r5_start + 1, 1, "Budget Utilized (Cost)", th_fmt)
            ws2.write(r5_start + 1, 2, "No. of Shipments", th_fmt)

            r5 = r5_start + 2
            top_cust = t5_customer.head(15)
            for _, row in top_cust.iterrows():
                ws2.write(r5, 0, row['Customer'], cell_txt)
                ws2.write(r5, 1, row['Budget_Utilized'], cell_curr)
                ws2.write(r5, 2, row['No_of_Shipments'], cell_num)
                r5 += 1
            ws2.write(r5, 0, "Grand Total (Top Customers)", total_txt)
            ws2.write_formula(r5, 1, f"=SUM(B{r5_start + 3}:B{r5})", total_curr)
            ws2.write_formula(r5, 2, f"=SUM(C{r5_start + 3}:C{r5})", total_num)

            # Chart 3: Customer Bar Chart
            chart3 = workbook.add_chart({'type': 'bar'})
            chart3.add_series({
                'name': 'Budget Utilized',
                'categories': f"='Page 2 - Region & Customer'!$A${r5_start + 3}:$A${r5}",
                'values': f"='Page 2 - Region & Customer'!$B${r5_start + 3}:$B${r5}",
                'fill': {'color': '#00A0DC'},
            })
            chart3.set_title({'name': 'Top Customers by Courier Spend'})
            chart3.set_x_axis({'name': 'Budget Utilized (PKR)'})
            chart3.set_size({'width': 480, 'height': 380})
            ws2.insert_chart(f'E{r5_start + 1}', chart3)

            ws2.set_column('A:A', 35)
            ws2.set_column('B:B', 22)
            ws2.set_column('C:C', 16)
            ws2.set_column('D:D', 4)

            workbook.close()

            # Download Button
            file_title = "_".join(selected_months)
            st.download_button(
                label="📥 Download Executive Excel Report (.xlsx)",
                data=excel_buffer.getvalue(),
                file_name=f"Courier_Expense_Report_{file_title}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error reading report: {e}")
