import os.path
import pathlib

from pdfreportexporter.excelreportexporter.testreportexporter import TestReportExporter
from pdfreportexporter.exceltopdf.exceltopdf import ExcelToPdf


def main():
    """
    テンプレート(Excelファイル)からPDF形式の帳票を出力する
    :return:
    """

    # テンプレートから帳票(Excel)を作成
    report_exporter = TestReportExporter()
    report_exporter.export(os.path.join("output", "test-report-1.xlsx"))

    # 帳票をExcel形式からPDF形式に変換
    excel_to_pdf = ExcelToPdf()
    excel_filepath = pathlib.Path(os.path.join("output", "test-report-1.xlsx"))
    pdf_filepath = pathlib.Path(os.path.join("output", "test-report-1.pdf"))
    excel_to_pdf.convert(str(excel_filepath.resolve()), "Sheet1", str(pdf_filepath.resolve()))


if __name__ == '__main__':
    main()
