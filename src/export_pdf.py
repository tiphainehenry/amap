import pandas as pd
from jinja2 import Template
import pdfkit
from pathlib import Path
from datetime import datetime, timedelta

HTML_TEMPLATE = """
<html>
<head>
  <meta charset="utf-8">
<style>
    table {
        border-collapse: collapse;
        width: 100%;
        font-family: Arial, sans-serif;
    }
    td, th {
        border: none;
        padding: 4px;
        text-align: left;
    }
</style>
</head>
<body>
  {% for title, table_html in tables %}
    <h2>{{ title }}</h2>
    {{ table_html|safe }}
  {% endfor %}
</body>
</html>
"""

def get_next_saturday_sheetname():
    today = datetime.today()
    days_until_saturday = (5 - today.weekday()) % 7
    saturday = today + timedelta(days=days_until_saturday)
    return saturday.strftime("%Y-%m-%d")

import pandas as pd
import pdfkit
from pathlib import Path

def generate_pdf(excel_path, sheet_names, pdf_output):
    config = pdfkit.configuration(wkhtmltopdf=r"C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe")

    STYLE = """
    <style>
        table {
            border-collapse: collapse;
            width: 100%;
            font-family: Arial, sans-serif;
            font-size: 12px;
        }
        td, th {
            padding: 4px;
            text-align: left;
        }
        tr.name-row td {
            border: 1px solid #000;
        }
        tr.other-row td {
            border: none;
        }
    </style>
    """

    def df_to_html_custom(df: pd.DataFrame) -> str:
        df.fillna("", inplace=True)
        html_rows = []

        for _, row in df.iterrows():
            values = row.astype(str).tolist()
            first_col = values[0].strip().lower()
            is_name_row = bool(values[0].strip()) and bool(values[1].strip())
            is_cumul_row = first_col == "cumul"
            cls = "name-row" if is_name_row or is_cumul_row else "other-row"
            row_html = f'<tr class="{cls}">' + "".join(f"<td>{v}</td>" for v in values) + "</tr>"
            html_rows.append(row_html)

        return "<table>\n" + "\n".join(html_rows) + "\n</table>"

    # 📄 Display names for sheets
    display_names = {
        "legumes_merged": "Légumes",
        "oeufs_merged": "Œufs"
    }

    html_parts = []
    xls = pd.ExcelFile(excel_path)

    for i, sheet in enumerate(sheet_names):
        df = xls.parse(sheet, header=None)
        html_table = df_to_html_custom(df)
        page_break = '<div style="page-break-before: always;"></div>' if i > 0 else ""
        title = {"legumes_merged": "Légumes", "oeufs_merged": "Œufs"}.get(sheet, sheet)

        html_parts.append(
            f"""{page_break}
    <h2>{title}</h2>
    {STYLE}
    {html_table}
    """
        )

    # Ensure full UTF-8 compatibility
    tmp_html_path = Path("tmp_combined.html")
    tmp_html_path.write_text(
        "<!DOCTYPE html><html><head><meta charset='UTF-8'></head><body>" +
        "\n".join(html_parts) +
        "</body></html>", encoding="utf-8"
    )


    pdfkit.from_file(str(tmp_html_path), str(pdf_output), configuration=config)
    tmp_html_path.unlink(missing_ok=True)

def main():
    date_str = get_next_saturday_sheetname()
    excel_file = Path(f"distrib_amap_{date_str}.xlsx")
    pdf_file = Path(f"distrib_amap_{date_str}.pdf")
    sheets = ["legumes_merged", "oeufs_merged"]

    if excel_file.exists():
        generate_pdf(excel_file, sheets, pdf_file)
    else:
        print(f"❌ Excel file not found: {excel_file}")

if __name__ == "__main__":
    main()
