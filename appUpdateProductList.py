import streamlit as st
import pandas as pd
import openpyxl
from copy import copy
import io

st.set_page_config(page_title="Product List Sync + Style", layout="wide")
st.title("📦 Product List Synchronizer")

def get_data_start(sheet):
    for row in sheet.iter_rows(min_row=1, max_row=25):
        cell_val = str(row[1].value).strip() if row[1].value else ""
        if cell_val.startswith('BP') or cell_val.startswith('100'):
            return row[0].row
    return 7

master_file = st.file_uploader("1. Upload Master List", type=['xlsx'])
daily_file = st.file_uploader("2. Upload Daily Sheet", type=['xlsx'])

if master_file and daily_file:
    df_master_raw = pd.read_excel(master_file, header=None)
    master_data = []
    for _, row in df_master_raw.iterrows():
        code, name = str(row[0]).strip(), str(row[1]).strip()
        if code.lower() not in ['nan', 'none', '']:
            master_data.append({'Item Code': code, 'Master_Name': name})
    df_master = pd.DataFrame(master_data).drop_duplicates('Item Code')

    wb = openpyxl.load_workbook(daily_file)
    selected_sheet = st.selectbox("Select Sheet:", wb.sheetnames)

    if st.button("Sync & Preserve Formatting"):
        sheet = wb[selected_sheet]
        data_start = get_data_start(sheet)
        
        # 1. Capture Data AND Styles
        original_rows_data = {}
        # We also capture a 'Generic Style' from the first data row to use for NEW items
        default_style_row = [] 

        for row in sheet.iter_rows(min_row=data_start):
            item_code = str(row[1].value).strip()
            row_styles = []
            for cell in row:
                row_styles.append({
                    'value': cell.value,
                    'fill': copy(cell.fill),
                    'font': copy(cell.font),
                    'border': copy(cell.border),
                    'alignment': copy(cell.alignment),
                    'number_format': cell.number_format
                })
            
            if not default_style_row:
                default_style_row = row_styles

            if item_code not in original_rows_data:
                original_rows_data[item_code] = row_styles

        # 2. Build the New Order
        final_list_of_row_styles = []
        for m_code in df_master['Item Code']:
            if m_code in original_rows_data:
                final_list_of_row_styles.append(original_rows_data[m_code])
            else:
                # NEW ITEM: Use default styles but change the Code and Name
                new_row_style = [copy(s) for s in default_style_row]
                # Reset values for the new row
                for i in range(len(new_row_style)):
                    new_row_style[i]['value'] = None
                
                new_row_style[0]['value'] = m_code # Col A
                new_row_style[1]['value'] = m_code # Col B
                m_name = df_master[df_master['Item Code'] == m_code]['Master_Name'].values[0]
                new_row_style[3]['value'] = m_name # Col D
                final_list_of_row_styles.append(new_row_style)

        # 3. Add items only in Daily
        for d_code, d_styles in original_rows_data.items():
            if d_code not in df_master['Item Code'].values:
                final_list_of_row_styles.append(d_styles)

        # 4. Wipe and Rewrite
        # Instead of deleting, we clear values to avoid breaking Excel's internal style table
        for row in sheet.iter_rows(min_row=data_start):
            for cell in row:
                cell.value = None

        for r_idx, styled_row in enumerate(final_list_of_row_styles, start=data_start):
            for c_idx, s in enumerate(styled_row, start=1):
                cell = sheet.cell(row=r_idx, column=c_idx)
                
                # Apply Value
                val = s['value']
                cell.value = None if pd.isna(val) or str(val).lower() in ['nan', 'none'] else val
                
                # Apply Styles
                cell.fill = s['fill']
                cell.font = s['font']
                cell.border = s['border']
                cell.alignment = s['alignment']
                cell.number_format = s['number_format']

        output = io.BytesIO()
        wb.save(output)
        st.success("Sync complete! Check the file for correct highlighting.")
        st.download_button("Download Updated File", output.getvalue(), f"Fixed_Styled_{selected_sheet}.xlsx")