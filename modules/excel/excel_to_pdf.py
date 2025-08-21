import pythoncom
import win32com.client


def excel_to_pdf_converter(excel_path, pdf_path):
    pythoncom.CoInitialize()  # Ensure COM is initialized
    excel = None
    wb = None

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(excel_path)

        # Setup page layout for all sheets
        for sheet in wb.Worksheets:
            sheet.PageSetup.Zoom = False
            sheet.PageSetup.FitToPagesWide = 1
            sheet.PageSetup.FitToPagesTall = False

        wb.ExportAsFixedFormat(0, pdf_path)

    except Exception as e:
        print(f"Error occurred: {e}")
        if wb:
            wb.Close(False)
        if excel:
            excel.Quit()

        # Re-initialize for fallback (not strictly needed here anymore)
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(excel_path)
        wb.Worksheets[0].ExportAsFixedFormat(0, pdf_path)

    finally:
        if wb:
            wb.Close(False)
        if excel:
            excel.Quit()
        pythoncom.CoUninitialize()
