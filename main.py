from openpyxl import load_workbook
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from datetime import datetime

def read_data_sheet(path: str, sheet_name: str = "DATA"):
    """Read the evaluated values (not formulas) from the DATA sheet."""
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet_name]

    # Get column headers
    data = list(ws.values)
    columns = data[0]
    df = pd.DataFrame(data[1:], columns=columns)

    # Extract column widths (approximate points)
    col_widths = []
    for col in ws.column_dimensions.values():
        width = col.width if col.width else 10
        # Excel width ~ 1 char = 7 points approx
        col_widths.append(width * 7)
    # If there are fewer column_dimensions entries than columns
    while len(col_widths) < len(df.columns):
        col_widths.append(70)

    return df, col_widths

def format_value(value):
    """Convert Excel-like values to nicely formatted strings."""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float):
        # Decide whether to show decimals
        return f"{value:,.2f}" if abs(value) < 1e7 else f"{value:,.0f}"
    return str(value) if value is not None else ""

def dataframe_to_styled_pdf(df: pd.DataFrame, pdf_path: str, col_widths=None):
    """Export DataFrame to PDF with Excel-like formatting."""
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(pdf_path, pagesize=landscape(A4), leftMargin=20, rightMargin=20)
    elements = []

    # Title
    elements.append(Paragraph("DATA Sheet Export", styles["Heading2"]))
    elements.append(Spacer(1, 12))

    # Format data
    data = [df.columns.tolist()] + [[format_value(v) for v in row] for row in df.values]

    # Create table with column widths
    table = Table(data, repeatRows=1, colWidths=col_widths)

    # Table style similar to Excel
    style = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4F81BD")),  # header background
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
    ])

    # Alternate row shading
    for i in range(1, len(data)):
        if i % 2 == 0:
            style.add("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F2F2F2"))

    table.setStyle(style)
    elements.append(table)
    doc.build(elements)

    print(f"✅ PDF successfully generated at: {pdf_path}")


if __name__ == "__main__":
    xlsm_path = "your_file.xlsm"
    pdf_path = "DATA_export.pdf"

    df, widths = read_data_sheet(xlsm_path, "DATA")
    dataframe_to_styled_pdf(df, pdf_path, col_widths=widths)
