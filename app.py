import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import zipfile
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go

# Page Configuration
st.set_page_config(
    page_title="Courier Intelligence",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Modern UI/UX CSS Injection
st.markdown("""
    <style>
    /* Global App Background */
    .stApp {
        background-color: #F3F4F6;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    /* Elegant Header */
    .hero-header {
        background: linear-gradient(100deg, #0f172a 0%, #1e3a8a 50%, #0284c7 100%);
        padding: 30px 40px;
        border-radius: 16px;
        color: white;
        margin-bottom: 30px;
        box-shadow: 0 10px 25px rgba(15, 23, 42, 0.1);
    }
    .hero-header h1 {
        color: #FFFFFF !important;
        margin: 0;
        font-size: 32px;
        font-weight: 800;
        letter-spacing: -1px;
    }
    .hero-header p {
        color: #94A3B8 !important;
        margin: 8px 0 0 0;
        font-size: 15px;
        font-weight: 400;
    }
    
    /* Modern KPI Cards */
    .kpi-container {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 24px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        transition: transform 0.2s ease;
    }
    .kpi-container:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.08);
    }
    .kpi-title {
        font-size: 13px;
        font-weight: 600;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        margin-bottom: 8px;
    }
    .kpi-metric {
        font-size: 28px;
        font-weight: 800;
        color: #0F172A;
    }
    .kpi-good {
        color: #059669;
    }
    
    /* Clean Sidebar */
    [data-testid="stSidebar"] {
        background-color: #FFFFFF !important;
        border-right: 1px solid #E2E8F0;
    }
    
    /* Download Button Upgrade */
    .stDownloadButton button {
        background: linear-gradient(to right, #0284c7, #0369a1) !important;
        color: white !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
        padding: 12px 28px !important;
        border: none !important;
        box-shadow: 0 4px 12px rgba(2, 132, 199, 0.25) !important;
        transition: all 0.3s ease !important;
    }
    .stDownloadButton button:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 6px 16px rgba(2, 132, 199, 0.35) !important;
    }
    </style>
""", unsafe_allow_html=True)

# Main Application Header
st.markdown("""
    <div class="hero-header">
        <h1>📊 Courier Spend Analytics</h1>
        <p>Upload your monthly detail report to generate an interactive dashboard and export presentation-ready workbooks.</p>
    </div>
""", unsafe_allow_html=True)

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

with st.sidebar:
    st.markdown("<div style='text-align: center; font-size: 60px; margin-bottom: 10px;'>🚀</div>", unsafe_allow_html=True)
    st.markdown("<h3 style='text-align: center; color: #0F172A;'>Data Import</h3>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Upload Detail Report (.xlsx)", type=["xlsx"])
    st.markdown("---")

if uploaded_file is not None:
    try:
        raw_df = parse_excel_safely(uploaded_file)
        raw_df = raw_df.dropna(subset=['Sales Executive']).copy()
        raw_df['Budget Utilized'] = pd.to_numeric(raw_df['Budget Utilized'], errors='coerce').fillna(0)
        raw_df['Monthly Budget'] = pd.to_numeric(raw_df['Monthly Budget'], errors='coerce').fillna(0)

        available_months = raw_df['Month'].dropna().unique().tolist()
        with st.sidebar:
            st.subheader("Filter Settings")
            selected_months = st.multiselect("Select Period:", options=available_months, default=available_months)

        if not selected_months:
            st.info("👈 Please select a month from the sidebar to view data.")
        else:
            df = raw_df[raw_df['Month'].isin(selected_months)]

            total_spend = df['Budget Utilized'].sum()
            total_budget = df.groupby('Sales Executive')['Monthly Budget'].first().sum()
            total_variance = total_budget - total_spend
            total_shipments = len(df)

            k1, k2, k3, k4 = st.columns(4)
            with k1:
                st.markdown(f"""
                    <div class="kpi-container">
                        <div class="kpi-title">Total Spend</div>
                        <div class="kpi-metric">₨ {total_spend:,.0f}</div>
                    </div>
                """, unsafe_allow_html=True)
            with k2:
                st.markdown(f"""
                    <div class="kpi-container">
                        <div class="kpi-title">Allocated Budget</div>
                        <div class="kpi-metric">₨ {total_budget:,.0f}</div>
                    </div>
                """, unsafe_allow_html=True)
            with k3:
                st.markdown(f"""
                    <div class="kpi-container">
                        <div class="kpi-title">Budget Variance</div>
                        <div class="kpi-metric kpi-good">₨ {total_variance:,.0f}</div>
                    </div>
                """, unsafe_allow_html=True)
            with k4:
                st.markdown(f"""
                    <div class="kpi-container">
                        <div class="kpi-title">Total Shipments</div>
                        <div class="kpi-metric">{total_shipments:,}</div>
                    </div>
                """, unsafe_allow_html=True)

            st.write("<br>", unsafe_allow_html=True)

            # Aggregations
            t1_exec = df.groupby('Sales Executive')['Budget Utilized'].sum().reset_index().sort_values('Budget Utilized', ascending=False)
            t2_shipments = df.groupby('Sales Executive').size().reset_index(name='Couriers').sort_values('Couriers', ascending=False)
            t3_budget = df.groupby('Sales Executive').agg(
                Budget_Allocated=('Monthly Budget', 'first'),
                Budget_Utilized=('Budget Utilized', 'sum')
            ).reset_index().sort_values('Budget_Allocated', ascending=False)
            t3_budget['Variance'] = t3_budget['Budget_Allocated'] - t3_budget['Budget_Utilized']

            t4_region = df.groupby('Country').agg(
                Budget_Utilized=('Budget Utilized', 'sum'),
                No_of_Shipments=('Country', 'count')
            ).reset_index().sort_values('Budget_Utilized', ascending=False)

            t5_customer = df.groupby('Customer').agg(
                Budget_Utilized=('Budget Utilized', 'sum'),
                No_of_Shipments=('Customer', 'count')
            ).reset_index().sort_values('Budget_Utilized', ascending=False)

            tab1, tab2 = st.tabs(["💼 Executives Overview", "🌍 Regional Analytics"])

            with tab1:
                st.write("<br>", unsafe_allow_html=True)
                col_left, col_right = st.columns([1, 1.5])
                
                with col_left:
                    st.markdown("#### Budget Utilization Table")
                    st.dataframe(
                        t3_budget.rename(columns={'Budget_Allocated': 'Allocated', 'Budget_Utilized': 'Utilized', 'Variance': 'Surplus'}),
                        column_config={
                            "Allocated": st.column_config.NumberColumn(format="PKR %,.0f"),
                            "Utilized": st.column_config.NumberColumn(format="PKR %,.0f"),
                            "Surplus": st.column_config.NumberColumn(format="PKR %,.0f")
                        },
                        use_container_width=True, hide_index=True
                    )
                
                with col_right:
                    st.markdown("#### Executive Budget Comparison")
                    # Plotly Grouped Bar Chart
                    fig_exec = px.bar(
                        t3_budget, 
                        x='Sales Executive', 
                        y=['Budget_Allocated', 'Budget_Utilized'],
                        barmode='group',
                        color_discrete_map={'Budget_Allocated': '#0F172A', 'Budget_Utilized': '#0284C7'},
                        labels={'value': 'Amount (PKR)', 'variable': 'Budget Type'}
                    )
                    fig_exec.update_layout(
                        plot_bgcolor='rgba(0,0,0,0)',
                        paper_bgcolor='rgba(0,0,0,0)',
                        margin=dict(l=0, r=0, t=20, b=0),
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                    )
                    st.plotly_chart(fig_exec, use_container_width=True)

            with tab2:
                st.write("<br>", unsafe_allow_html=True)
                r_col1, r_col2 = st.columns(2)
                
                with r_col1:
                    st.markdown("#### Top Destinations")
                    fig_reg = px.bar(
                        t4_region.head(8), 
                        x='Budget_Utilized', 
                        y='Country',
                        orientation='h',
                        color_discrete_sequence=['#1E3A8A']
                    )
                    fig_reg.update_layout(
                        yaxis={'categoryorder':'total ascending'},
                        plot_bgcolor='rgba(0,0,0,0)',
                        xaxis_title="Spend (PKR)",
                        yaxis_title="",
                        margin=dict(l=0, r=0, t=10, b=0)
                    )
                    st.plotly_chart(fig_reg, use_container_width=True)

                with r_col2:
                    st.markdown("#### Top Customers")
                    fig_cust = px.bar(
                        t5_customer.head(8), 
                        x='Budget_Utilized', 
                        y='Customer',
                        orientation='h',
                        color_discrete_sequence=['#0284C7']
                    )
                    fig_cust.update_layout(
                        yaxis={'categoryorder':'total ascending'},
                        plot_bgcolor='rgba(0,0,0,0)',
                        xaxis_title="Spend (PKR)",
                        yaxis_title="",
                        margin=dict(l=0, r=0, t=10, b=0)
                    )
                    st.plotly_chart(fig_cust, use_container_width=True)

            st.write("<br>", unsafe_allow_html=True)
            st.markdown("---")
            
            # EXCEL EXPORT ENGINE (REMAINS INTACT FOR NATIVE EXCEL CHARTS)
            excel_buffer = BytesIO()
            workbook = pd.ExcelWriter(excel_buffer, engine='xlsxwriter').book

            ws1 = workbook.add_worksheet('Page 1 - Executive Summary')
            ws1.set_paper(9); ws1.set_portrait(); ws1.fit_to_pages(1, 0)
            ws2 = workbook.add_worksheet('Page 2 - Region & Customer')
            ws2.set_paper(9); ws2.set_portrait(); ws2.fit_to_pages(1, 0)

            title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#1F497D'})
            sec_hdr_fmt = workbook.add_format({'bold': True, 'font_size': 11, 'bg_color': '#D9E1F2', 'font_color': '#1F497D', 'border': 1})
            th_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F497D', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            cell_txt = workbook.add_format({'border': 1, 'valign': 'vcenter'})
            cell_num = workbook.add_format({'border': 1, 'num_format': '#,##0', 'align': 'right', 'valign': 'vcenter'})
            cell_curr = workbook.add_format({'border': 1, 'num_format': 'PKR #,##0', 'align': 'right', 'valign': 'vcenter'})
            total_txt = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#F2F2F2', 'valign': 'vcenter'})
            total_curr = workbook.add_format({'bold': True, 'border': 1, 'num_format': 'PKR #,##0', 'bg_color': '#F2F2F2', 'align': 'right', 'valign': 'vcenter'})
            total_num = workbook.add_format({'bold': True, 'border': 1, 'num_format': '#,##0', 'bg_color': '#F2F2F2', 'align': 'right', 'valign': 'vcenter'})

            month_label = ", ".join(selected_months)
            ws1.write(0, 0, f"Courier Budget Detail Report - {month_label}", title_fmt)

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
            chart1.set_title({'name': 'Executive Budget Comparison'})
            chart1.set_x_axis({'name': 'Executive'})
            chart1.set_y_axis({'name': 'Amount (PKR)', 'major_gridlines': {'visible': True}})
            chart1.set_size({'width': 540, 'height': 320})
            ws1.insert_chart('E13', chart1)

            ws1.set_column('A:A', 22); ws1.set_column('B:C', 16); ws1.set_column('D:D', 4)
            ws1.set_column('E:E', 22); ws1.set_column('F:H', 18)

            ws2.write(0, 0, f"Courier Spend by Region & Customer - {month_label}", title_fmt)

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

            chart2 = workbook.add_chart({'type': 'bar'})
            chart2.add_series({
                'name': 'Budget Utilized',
                'categories': f"='Page 2 - Region & Customer'!$A$5:$A${r4}",
                'values': f"='Page 2 - Region & Customer'!$B$5:$B${r4}",
                'fill': {'color': '#1F497D'},
            })
            chart2.set_title({'name': 'Top Destinations'})
            chart2.set_size({'width': 480, 'height': 280})
            ws2.insert_chart('E3', chart2)

            r5_start = max(r4 + 3, 18)
            ws2.write(r5_start, 0, "Table 5: Customer Wise", sec_hdr_fmt)
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
            ws2.write(r5, 0, "Grand Total", total_txt)
            ws2.write_formula(r5, 1, f"=SUM(B{r5_start + 3}:B{r5})", total_curr)
            ws2.write_formula(r5, 2, f"=SUM(C{r5_start + 3}:C{r5})", total_num)

            chart3 = workbook.add_chart({'type': 'bar'})
            chart3.add_series({
                'name': 'Budget Utilized',
                'categories': f"='Page 2 - Region & Customer'!$A${r5_start + 3}:$A${r5}",
                'values': f"='Page 2 - Region & Customer'!$B${r5_start + 3}:$B${r5}",
                'fill': {'color': '#00A0DC'},
            })
            chart3.set_title({'name': 'Top Customers'})
            chart3.set_size({'width': 480, 'height': 380})
            ws2.insert_chart(f'E{r5_start + 1}', chart3)

            ws2.set_column('A:A', 35); ws2.set_column('B:B', 22); ws2.set_column('C:C', 16); ws2.set_column('D:D', 4)

            workbook.close()

            file_title = "_".join(selected_months)
            
            # Action Button
            st.markdown("### 📥 Generate Export")
            st.download_button(
                label="Download Formatted Management Workbook (.xlsx)",
                data=excel_buffer.getvalue(),
                file_name=f"Courier_Expense_Report_{file_title}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error processing report: {e}")
