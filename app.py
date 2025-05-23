# T12 Formatter Streamlit App
# Final version with revised steps and corrected freeze pane position

import streamlit as st
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from datetime import datetime
import tempfile

def format_t12(file_path):
    wb = load_workbook(file_path)
    ws = wb.active

    # 1. Unmerge cells in A1:N60
    for merged_range in list(ws.merged_cells.ranges):
        if merged_range.max_row <= 60 and merged_range.max_col <= 14:
            ws.unmerge_cells(str(merged_range))

    # 2. Align Left cells A6-A8
    for row in range(6, 9):
        ws[f"A{row}"].alignment = Alignment(horizontal='left')

    # 3 & 11. Set Column Widths B–N to 12 pixels
    for col in range(2, 15):
        ws.column_dimensions[get_column_letter(col)].width = 12

    # 4. Shade A45:N45
    fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    for col in range(1, 15):
        ws.cell(row=45, column=col).fill = fill

    # 5. Bold A21:N21, A25:N25, A43:N43, A45:N45, A55:N55
    bold_font = Font(bold=True)
    for row in [21, 25, 43, 45, 55]:
        for col in range(1, 15):
            ws.cell(row=row, column=col).font = bold_font

    # 6. Delete rows 1–5, 9–11, and 60
    for row in sorted([1, 2, 3, 4, 5, 9, 10, 11, 60], reverse=True):
        ws.delete_rows(row)

    # 7. Freeze Pane at B6
    ws.freeze_panes = "B6"

    # 8. Use A3 for filename
    date_value = ws["A3"].value if ws["A3"].value else "NoDate"
    try:
        parsed_date = datetime.strptime(str(date_value), "%B %d, %Y")
        formatted_date = parsed_date.strftime("%Y-%m")
    except ValueError:
        formatted_date = "Unknown_Date"

    # 9 & 10. Set row heights
    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 16

    # 12. Add "Total" to N5 and right-align
    ws["N5"].value = "Total"
    ws["N5"].alignment = Alignment(horizontal="right")

    # 13. Hide gridlines
    ws.sheet_view.showGridLines = False

    # 14. Save output
    property_name = ws["A1"].value or "Property"
    safe_property = str(property_name).replace(" ", "_").replace("/", "-").strip()
    filename = f"{safe_property}_T12_{formatted_date}.xlsx"
    output_path = os.path.join(tempfile.gettempdir(), filename)
    wb.save(output_path)

    return output_path

# Streamlit UI
st.title("📊 T12 Formatter for Multifamily Reports")
st.write("Upload a raw T12 Excel file and get a perfectly formatted version for investors and lenders.")

uploaded_file = st.file_uploader("Upload your T12 Excel file", type=["xlsx"])

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.read())
        formatted_path = format_t12(tmp.name)

    with open(formatted_path, "rb") as f:
        st.download_button(
            label="📥 Download Formatted T12",
            data=f,
            file_name=os.path.basename(formatted_path),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
