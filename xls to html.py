import os
import base64
import tempfile
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox

import win32com.client as win32


def create_html_from_pdf(pdf_path, html_path, title):
    with open(pdf_path, "rb") as f:
        pdf_base64 = base64.b64encode(f.read()).decode("utf-8")

    html = f"""<!DOCTYPE html>
<html lang="tr">
<head>
<meta charset="UTF-8">
<title>{title}</title>
<style>
body {{
    margin: 0;
    background: #f3f4f6;
    font-family: Arial, sans-serif;
}}

header {{
    background: #1f2937;
    color: white;
    padding: 12px 20px;
    font-size: 18px;
    font-weight: bold;
}}

.viewer {{
    width: 100%;
    height: calc(100vh - 52px);
}}

iframe {{
    width: 100%;
    height: 100%;
    border: none;
    background: white;
}}
</style>
</head>
<body>
<header>{title}</header>
<div class="viewer">
    <iframe src="data:application/pdf;base64,{pdf_base64}"></iframe>
</div>
</body>
</html>
"""

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)


def process_excel_file(excel_path, output_folder):
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False

    file_name = os.path.splitext(os.path.basename(excel_path))[0]
    html_path = os.path.join(output_folder, f"{file_name}.html")

    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    temp_pdf_path = temp_pdf.name
    temp_pdf.close()

    try:
        wb = excel.Workbooks.Open(
            excel_path,
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True
        )

        printable_sheets = []

        for ws in wb.Worksheets:
            if ws.Visible != -1:
                continue

            print_area = ws.PageSetup.PrintArea

            if print_area and str(print_area).strip():
                printable_sheets.append(ws.Name)

        if not printable_sheets:
            wb.Close(SaveChanges=False)
            messagebox.showwarning(
                "Yazdırma alanı bulunamadı",
                f"{os.path.basename(excel_path)} içinde yazdırma alanı tanımlı görünür sekme bulunamadı."
            )
            return None

        wb.Worksheets(printable_sheets).Select()

        excel.ActiveSheet.ExportAsFixedFormat(
            Type=0,
            Filename=temp_pdf_path,
            Quality=0,
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False
        )

        wb.Close(SaveChanges=False)

        create_html_from_pdf(
            pdf_path=temp_pdf_path,
            html_path=html_path,
            title=file_name
        )

        return html_path

    finally:
        try:
            excel.Quit()
        except Exception:
            pass

        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)


def main():
    root = tk.Tk()
    root.withdraw()

    excel_files = filedialog.askopenfilenames(
        title="HTML'e dönüştürülecek Excel dosyalarını seçin",
        filetypes=[
            ("Excel Dosyaları", "*.xlsx *.xlsm"),
            ("Tüm Dosyalar", "*.*")
        ]
    )

    if not excel_files:
        return

    output_folder = filedialog.askdirectory(
        title="HTML dosyalarının kaydedileceği klasörü seçin"
    )

    if not output_folder:
        return

    created = []

    for excel_file in excel_files:
        try:
            html_path = process_excel_file(excel_file, output_folder)

            if html_path:
                created.append(html_path)

        except Exception as e:
            messagebox.showerror(
                "Hata",
                f"Dosya işlenirken hata oluştu:\n\n"
                f"{excel_file}\n\n"
                f"{e}\n\n"
                f"{traceback.format_exc()}"
            )

    messagebox.showinfo(
        "İşlem tamamlandı",
        f"{len(created)} HTML dosyası oluşturuldu.\n\n"
        f"Yalnızca yazdırma alanı tanımlı sayfalar PDF'e aktarıldı.\n"
        f"PDF dosyaları geçici oluşturulup silindi."
    )


if __name__ == "__main__":
    main()