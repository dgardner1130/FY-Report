import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
import tempfile
import os
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment, numbers
from openpyxl.formatting.rule import CellIsRule

# UI
st.title("📊 FY Project Report Creator")
st.write("Upload a CSV file and specify the fiscal year start.")

uploaded_file = st.file_uploader("Upload Fiscal Year Data", type="csv")
fiscal_year_start_month = 7
fiscal_year_start_year = st.number_input("Fiscal Year Start Year", value=2024, step=1)

if uploaded_file:
    data = pd.read_csv(uploaded_file)
    data['Date Comment Letter Sent'] = pd.to_datetime(data['Date Comment Letter Sent'], errors='coerce')
    data['Date Submitted'] = pd.to_datetime(data['Date Submitted'], errors='coerce')

    fiscal_months = list(range(fiscal_year_start_month, 13)) + list(range(1, fiscal_year_start_month))
    months = {}

    for month in fiscal_months:
        year = fiscal_year_start_year if month >= fiscal_year_start_month else fiscal_year_start_year + 1
        start_date = pd.Timestamp(year, month, 1)
        end_date = pd.Timestamp(year+1, 1, 1) - pd.Timedelta(days=1) if month == 12 else pd.Timestamp(year, month+1, 1) - pd.Timedelta(days=1)

        month_df = data[(data['Date Comment Letter Sent'] >= start_date) & 
                        (data['Date Comment Letter Sent'] <= end_date)][
                            ['Date Submitted', 'Development Name', 'Project No', 
                             'Review Cycle - ENG', 'Review Cycle - SUR', 'Review Cycle - PLN',
                             'Date Comment Letter Sent']
                        ].copy()

        month_df['Length of Review'] = (month_df['Date Comment Letter Sent'] - month_df['Date Submitted']).dt.days
        key = f'{start_date.strftime("%b %Y")}'
        months[key] = month_df

    reviewValues = [df['Length of Review'].tolist() for df in months.values()]
    lenResults = [len(lst) for lst in reviewValues]
    countExceeds30 = [sum(1 for x in lst if x > 30) for lst in reviewValues]
    shortCount = [total - long for total, long in zip(lenResults, countExceeds30)]
    avgLength = [round(np.mean(lst), 2) if lst else 0 for lst in reviewValues]
    percentage = [round((total - long)/total * 100, 1) if total else 0 for total, long in zip(lenResults, countExceeds30)]

    summary_df = pd.DataFrame({
        'Month': list(months.keys()),
        'Total Reviews': lenResults,
        '> 30 Days': countExceeds30,
        'Average Review Length': avgLength,
        '% Reviews ≤ 30 Days': percentage
    })

    # --- Plots ---
    x = np.arange(len(months))

    # Stacked Bar
    fig1, ax1 = plt.subplots(figsize=(12, 6))
    ax1.bar(x, shortCount, label='≤ 30 Days')
    ax1.bar(x, countExceeds30, bottom=shortCount, label='> 30 Days')
    ax1.set_title('Review Duration Breakdown per Month')
    ax1.set_xticks(x)
    ax1.set_xticklabels(months.keys(), rotation=45)
    ax1.set_ylabel('Number of Projects')
    ax1.legend()
    st.pyplot(fig1)

    # Line Chart
    fig2, ax2 = plt.subplots(figsize=(12, 6))
    ax2.plot(x, avgLength, marker='o')
    ax2.set_title('Average Review Duration per Month')
    ax2.set_xticks(x)
    ax2.set_xticklabels(months.keys(), rotation=45)
    ax2.set_ylabel('Average Days')
    st.pyplot(fig2)

    # --- Export Button ---
    if st.button("📥 Export Excel Report"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            excel_path = tmp.name
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
                workbook = writer.book
                summary_ws = writer.sheets['Summary']

                # Format header
                for cell in summary_ws[1]:
                    cell.font = Font(size=12, bold=True)
                    cell.fill = PatternFill("solid", fgColor="b5dbaf")
                    cell.alignment = Alignment(horizontal = 'center', vertical = 'center', wrap_text = True)

                for col in summary_ws.columns:
                        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
                        col_letter = col[0].column_letter
                        summary_ws.column_dimensions[col_letter].width = max_length + 5

                # Charts
                img_buffer1 = BytesIO()
                fig1.savefig(img_buffer1, format='png')
                img_buffer1.seek(0)
                img1 = XLImage(img_buffer1)
                summary_ws.add_image(img1, "G1")

                img_buffer2 = BytesIO()
                fig2.savefig(img_buffer2, format='png')
                img_buffer2.seek(0)
                img2 = XLImage(img_buffer2)
                summary_ws.add_image(img2, "G31")

                # Write monthly sheets
                for name, df in months.items():
                    df.to_excel(writer, sheet_name=name[:31], index=False)
                    ws = writer.sheets[name[:31]]

                    for cell in ws[1]:
                        cell.font = Font(size=12, bold=True)
                        cell.fill = PatternFill("solid", fgColor="b5dbaf")
                        cell.alignment = Alignment(horizontal = 'center', vertical = 'center', wrap_text = True)

                    for col in ws.columns:
                        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
                        col_letter = col[0].column_letter
                        ws.column_dimensions[col_letter].width = max_length + 5

                    # Date formatting
                    for row in ws.iter_rows(min_row=2, min_col=1, max_col=2):
                        for cell in row:
                            cell.number_format = numbers.FORMAT_DATE_YYYYMMDD2
                    for row in ws.iter_rows(min_row=2, min_col=6, max_col=7):
                        for cell in row:
                            cell.number_format = numbers.FORMAT_DATE_YYYYMMDD2

                    # Conditional formatting
                    ws.conditional_formatting.add('H2:H1000', CellIsRule(operator = 'greaterThan', formula = ['30'], fill = PatternFill(start_color = 'FFC7CE', end_color = 'FFC7CE', fill_type = 'solid')))

            with open(excel_path, 'rb') as f:
                st.download_button(
                    label="📂 Download Excel Report",
                    data=f,
                    file_name="monthly_review_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
