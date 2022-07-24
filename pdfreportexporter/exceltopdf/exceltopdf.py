import win32com.client


class ExcelToPdf:

    @staticmethod
    def convert(excel_filepath, sheet_name, pdf_filepath):
        """
        ExcelファイルをPDFファイルに変換する
        :param excel_filepath: Excelファイルのパス
        :param sheet_name: PDFに変換するExcelファイルにおけるシートの名前
        :param pdf_filepath: PDFファイルのパス
        :return:
        """

        excel = win32com.client.Dispatch("Excel.Application")
        excel_file = excel.Workbooks.Open(excel_filepath)

        try:
            excel_file.WorkSheets(sheet_name).Activate()

            excel_file.ActiveSheet.ExportAsFixedFormat(0, pdf_filepath)
        finally:
            if excel_file is not None:
                excel_file.Close()
            if excel is not None:
                excel.Quit()
