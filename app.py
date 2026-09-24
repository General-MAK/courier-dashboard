import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import zipfile
from io import BytesIO
import plotly.express as px

# 1. PAGE CONFIGURATION
st.set_page_config(
    page_title="Courier Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. BULLETPROOF CSS OVERRIDE (FORCES LIGHT MODE & HIGH CONTRAST)
st.markdown("""
    <style>
    /* Force main background to light gray */
    .stApp, .main {
        background-color: #F8FAFC !important;
    }
    
    /* Force Sidebar background to solid white and text to black */
    [data-testid="stSidebar"] {
        background-color: #FFFFFF !important;
        border-right: 1px solid #E2E8F0 !important;
    }
    [data-testid="stSidebarNav"], [data-testid="stSidebar"] p, [data-testid="stSidebar"] span, [data-testid="stSidebar"] label, [data-testid="stSidebar"] div {
        color: #0F172A !important; 
    }

    /* Force all general text, headers, and labels to dark navy/black */
    h1, h2, h3, h4, h5, h6, p, label, .stMarkdown {
        color: #0F172A !important;
    }
    
    /* Header Banner styling */
    .main-header {
        background-color: #1E3A8A;
        padding: 25px 35px;
        border-radius: 12px;
        margin-bottom: 25px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .main-header h1 {
        color: #FFFFFF !important;
        margin: 0;
        font-size: 32px;
        font-weight: 800;
    }
    .main-header p {
        color: #E2E8F0 !important;
        margin-top: 5px;
        font-size: 16px;
    }
    
    /* KPI Cards */
    .kpi-card {
        background: #FFFFFF !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .kpi-title {
        font-size: 14px;
        font-weight: 700;
        color: #475569 !important;
        text-transform: uppercase;
    }
    .kpi-value {
        font-size: 28px;
        font-weight: 800;
        color: #0F172A !important;
        margin-top: 5px;
    }
    .kpi-good { color: #059669 !important; }
    
    /* Section Headers */
    .section-title {
        font-size: 20px;
        font-weight: 700;
        color: #1E3A8A !important;
        margin-top: 20px;
        margin-bottom: 10px;
        border-bottom: 2px solid #E2E8F0;
        padding-bottom: 5px;
    }

    /* Force Download Button to stand out (Bright Blue with White Text) */
    .stDownloadButton button {
        background-color: #0284C7 !important;
        border: none !important;
        border-radius: 6px !important;
        padding: 10px 20px !important;
    }
    .stDownloadButton button p, .stDownloadButton button span {
        color: #FFFFFF !important;
        font-weight: 700 !important;
        font-size: 16px !important;
    }
    </style>
""", unsafe_allow_html=True)

# 3. ROBUST EXCEL PARSER
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

# 4. APP HEADER & SIDEBAR
st.markdown("""
    <div class="main-header">
        <h1>📦 Executive Courier Dashboard</h1>
        <p>Upload data to view full analytics and download the 1-page Excel summary.</p>
    </div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.title("📂 Data Import")
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
            st.warning("Please select a month from the sidebar.")
        else:
            df = raw_df[raw_df['Month'].isin(selected_months)]

            # 5. DATA AGGREGATION
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
            ).reset_index().sort_values('Budget_Utilized', ascending=False).head(10)

            t5_customer = df.groupby('Customer').agg(
                Budget_Utilized=('Budget Utilized', 'sum'),
                No_of_Shipments=('Customer', 'count')
            ).reset_index().sort_values('Budget_Utilized', ascending=False).head(15)

            # KPIs
            total_spend = df['Budget Utilized'].sum()
            total_budget = df.groupby('Sales Executive')['Monthly Budget'].first().sum()
            total_variance = total_budget - total_spend
            total_shipments = len(df)

            k1, k2, k3, k4 = st.columns(4)
            with k1: st.markdown(f'<div class="kpi-card"><div class="kpi-title">Total Spend</div><div class="kpi-value">₨ {total_spend:,.0f}</div></div>', unsafe_allow_html=True)
            with k2: st.markdown(f'<div class="kpi-card"><div class="kpi-title">Allocated Budget</div><div class="kpi-value">₨ {total_budget:,.0f}</div></div>', unsafe_allow_html=True)
            with k3: st.markdown(f'<div class="kpi-card"><div class="kpi-title">Budget Variance</div><div class="kpi-value kpi-good">₨ {total_variance:,.0f}</div></div>', unsafe_allow_html=True)
            with k4: st.markdown(f'<div class="kpi-card"><div class="kpi-title">Total Shipments</div><div class="kpi-value">{total_shipments:,}</div></div>', unsafe_allow_html=True)

            st.write("<br>", unsafe_allow_html=True)

            # 6. VISIBLE LAYOUT
            
            # --- SECTION 1: EXECUTIVES ---
            st.markdown('<div class="section-title">1. Executive Budgets & Shipments</div>', unsafe_allow_html=True)
            col_a, col_b = st.columns([1, 1.5])
            with col_a:
                st.markdown("**Table 1: Expense by Executive**")
                st.dataframe(t1_exec.style.format({'Budget Utilized': 'PKR {:,.0f}'}), use_container_width=True, hide_index=True)
                st.markdown("**Table 2: Shipments by Executive**")
                st.dataframe(t2_shipments, use_container_width=True, hide_index=True)
            
            with col_b:
                st.markdown("**Table 3: Budget Utilization**")
                st.dataframe(t3_budget.style.format({'Budget_Allocated': 'PKR {:,.0f}', 'Budget_Utilized': 'PKR {:,.0f}', 'Variance': 'PKR {:,.0f}'}), use_container_width=True, hide_index=True)
                
                fig_exec = px.bar(t3_budget, x='Sales Executive', y=['Budget_Allocated', 'Budget_Utilized'], barmode='group',
                                  color_discrete_map={'Budget_Allocated': '#0F172A', 'Budget_Utilized': '#0284C7'})
                fig_exec.update_layout(title="Budget Comparison", margin=dict(t=30, b=0), plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig_exec, use_container_width=True)

            # --- SECTION 2: REGIONS ---
            st.markdown('<div class="section-title">2. Regional Spend (Top 10)</div>', unsafe_allow_html=True)
            col_c, col_d = st.columns([1, 1.5])
            with col_c:
                st.dataframe(t4_region.style.format({'Budget_Utilized': 'PKR {:,.0f}'}), use_container_width=True, hide_index=True)
            with col_d:
                fig_reg = px.bar(t4_region, x='Budget_Utilized', y='Country', orientation='h', color_discrete_sequence=['#1E3A8A'])
                fig_reg.update_layout(yaxis={'categoryorder':'total ascending'}, margin=dict(t=0, b=0), plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig_reg, use_container_width=True)

            # --- SECTION 3: CUSTOMERS ---
            st.markdown('<div class="section-title">3. Customer Spend (Top 15)</div>', unsafe_allow_html=True)
            col_e, col_f = st.columns([1, 1.5])
            with col_e:
                st.dataframe(t5_customer.style.format({'Budget_Utilized': 'PKR {:,.0f}'}), use_container_width=True, hide_index=True)
            with col_f:
                fig_cust = px.bar(t5_customer, x='Budget_Utilized', y='Customer', orientation='h', color_discrete_sequence=['#0284C7'])
                fig_cust.update_layout(yaxis={'categoryorder':'total ascending'}, margin=dict(t=0, b=0), plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig_cust, use_container_width=True)

            # 7. SINGLE-PAGE EXCEL GENERATION
            excel_buffer = BytesIO()
            workbook = pd.ExcelWriter(excel_buffer, engine='xlsxwriter').book
            
            ws = workbook.add_worksheet('Dashboard')
            ws.set_zoom(90)

            # Styles
            title_fmt = workbook.add_format({'bold': True, 'font_size': 16, 'font_color': '#1F497D'})
            sec_hdr_fmt = workbook.add_format({'bold': True, 'font_size': 12, 'bg_color': '#D9E1F2', 'font_color': '#1F497D', 'border': 1})
            th_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F497D', 'font_color': 'white', 'border': 1, 'align': 'center'})
            cell_txt = workbook.add_format({'border': 1})
            cell_num = workbook.add_format({'border': 1, 'num_format': '#,##0', 'align': 'right'})
            cell_curr = workbook.add_format({'border': 1, 'num_format': 'PKR #,##0', 'align': 'right'})
            total_txt = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#F2F2F2'})
            total_curr = workbook.add_format({'bold': True, 'border': 1, 'num_format': 'PKR #,##0', 'bg_color': '#F2F2F2', 'align': 'right'})
            total_num = workbook.add_format({'bold': True, 'border': 1, 'num_format': '#,##0', 'bg_color': '#F2F2F2', 'align': 'right'})

            ws.set_column('A:A', 35) 
            ws.set_column('B:B', 18) 
            ws.set_column('C:C', 15) 
            ws.set_column('D:D', 3)  
            ws.set_column('E:E', 25) 
            ws.set_column('F:H', 18) 

            month_label = ", ".join(selected_months)
            ws.write(0, 0, f"Executive Courier Dashboard - {month_label}", title_fmt)

            # TABLE 1
            ws.write(2, 0, "Table 1: Expense Summary", sec_hdr_fmt)
            ws.write(3, 0, "Executive", th_fmt); ws.write(3, 1, "Spend (PKR)", th_fmt); ws.write(3, 2, "Grand Total", th_fmt)
            r = 4
            for _, row in t1_exec.iterrows():
                ws.write(r, 0, row['Sales Executive'], cell_txt); ws.write(r, 1, row['Budget Utilized'], cell_curr); ws.write_formula(r, 2, f"=B{r+1}", cell_curr)
                r += 1
            ws.write(r, 0, "Grand Total", total_txt); ws.write_formula(r, 1, f"=SUM(B5:B{r})", total_curr); ws.write_formula(r, 2, f"=SUM(C5:C{r})", total_curr)

            # TABLE 2 
            r2_start = r + 2
            ws.write(r2_start, 0, "Table 2: Shipments", sec_hdr_fmt)
            ws.write(r2_start + 1, 0, "Executive", th_fmt); ws.write(r2_start + 1, 1, "Couriers", th_fmt)
            r2 = r2_start + 2
            for _, row in t2_shipments.iterrows():
                ws.write(r2, 0, row['Sales Executive'], cell_txt); ws.write(r2, 1, row['Couriers'], cell_num)
                r2 += 1
            ws.write(r2, 0, "Grand Total", total_txt); ws.write_formula(r2, 1, f"=SUM(B{r2_start + 3}:B{r2})", total_num)

            # TABLE 3 
            ws.write(2, 4, "Table 3: Budget Summary", sec_hdr_fmt)
            ws.write(3, 4, "Executive", th_fmt); ws.write(3, 5, "Allocated", th_fmt); ws.write(3, 6, "Utilized", th_fmt); ws.write(3, 7, "Variance", th_fmt)
            r3 = 4
            for _, row in t3_budget.iterrows():
                ws.write(r3, 4, row['Sales Executive'], cell_txt); ws.write(r3, 5, row['Budget_Allocated'], cell_curr)
                ws.write(r3, 6, row['Budget_Utilized'], cell_curr); ws.write_formula(r3, 7, f"=F{r3+1}-G{r3+1}", cell_curr)
                r3 += 1
            ws.write(r3, 4, "Grand Total", total_txt); ws.write_formula(r3, 5, f"=SUM(F5:F{r3})", total_curr)
            ws.write_formula(r3, 6, f"=SUM(G5:G{r3})", total_curr); ws.write_formula(r3, 7, f"=SUM(H5:H{r3})", total_curr)

            # CHART 1 
            chart1 = workbook.add_chart({'type': 'column'})
            chart1.add_series({'name': 'Allocated', 'categories': f'=Dashboard!$E$5:$E${r3}', 'values': f'=Dashboard!$F$5:$F${r3}', 'fill': {'color': '#1F497D'}})
            chart1.add_series({'name': 'Utilized', 'categories': f'=Dashboard!$E$5:$E${r3}', 'values': f'=Dashboard!$G$5:$G${r3}', 'fill': {'color': '#00A0DC'}})
            chart1.set_title({'name': 'Budget Comparison'})
            chart1.set_size({'width': 580, 'height': 280})
            ws.insert_chart(r3 + 2, 4, chart1)

            # TABLE 4 
            r4_start = max(r2 + 2, 22)
            ws.write(r4_start, 0, "Table 4: Top 10 Regions", sec_hdr_fmt)
            ws.write(r4_start + 1, 0, "Country", th_fmt); ws.write(r4_start + 1, 1, "Spend (PKR)", th_fmt); ws.write(r4_start + 1, 2, "Shipments", th_fmt)
            r4 = r4_start + 2
            for _, row in t4_region.iterrows():
                ws.write(r4, 0, row['Country'], cell_txt); ws.write(r4, 1, row['Budget_Utilized'], cell_curr); ws.write(r4, 2, row['No_of_Shipments'], cell_num)
                r4 += 1
            ws.write(r4, 0, "Grand Total", total_txt); ws.write_formula(r4, 1, f"=SUM(B{r4_start+3}:B{r4})", total_curr); ws.write_formula(r4, 2, f"=SUM(C{r4_start+3}:C{r4})", total_num)

            # CHART 2 
            chart2 = workbook.add_chart({'type': 'bar'})
            chart2.add_series({'name': 'Spend', 'categories': f'=Dashboard!$A${r4_start+3}:$A${r4}', 'values': f'=Dashboard!$B${r4_start+3}:$B${r4}', 'fill': {'color': '#1F497D'}})
            chart2.set_title({'name': 'Spend by Region'})
            chart2.set_size({'width': 580, 'height': 280})
            ws.insert_chart(r4_start, 4, chart2)

            # TABLE 5 
            r5_start = r4 + 2
            ws.write(r5_start, 0, "Table 5: Top 15 Customers", sec_hdr_fmt)
            ws.write(r5_start + 1, 0, "Customer", th_fmt); ws.write(r5_start + 1, 1, "Spend (PKR)", th_fmt); ws.write(r5_start + 1, 2, "Shipments", th_fmt)
            r5 = r5_start + 2
            for _, row in t5_customer.iterrows():
                ws.write(r5, 0, row['Customer'], cell_txt); ws.write(r5, 1, row['Budget_Utilized'], cell_curr); ws.write(r5, 2, row['No_of_Shipments'], cell_num)
                r5 += 1
            ws.write(r5, 0, "Grand Total", total_txt); ws.write_formula(r5, 1, f"=SUM(B{r5_start+3}:B{r5})", total_curr); ws.write_formula(r5, 2, f"=SUM(C{r5_start+3}:C{r5})", total_num)

            # CHART 3 
            chart3 = workbook.add_chart({'type': 'bar'})
            chart3.add_series({'name': 'Spend', 'categories': f'=Dashboard!$A${r5_start+3}:$A${r5}', 'values': f'=Dashboard!$B${r5_start+3}:$B${r5}', 'fill': {'color': '#00A0DC'}})
            chart3.set_title({'name': 'Spend by Customer'})
            chart3.set_size({'width': 580, 'height': 380})
            ws.insert_chart(r5_start, 4, chart3)

            workbook.close()

            st.markdown("---")
            st.markdown("### 📥 Download Export")
            st.download_button(
                label="Download 1-Page Excel Dashboard (.xlsx)",
                data=excel_buffer.getvalue(),
                file_name=f"Courier_Dashboard_{'_'.join(selected_months)}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error processing data: {e}")
