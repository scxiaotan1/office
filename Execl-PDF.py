import os
import win32com.client

def excel_to_pdf(input_file, output_file):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(input_file)
    wb.SaveAs(output_file, FileFormat=57)  # 57 represents PDF format
    wb.Close()

    excel.Quit()

def batch_convert_excel_to_pdf(input_folder, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for file_name in os.listdir(input_folder):
        if file_name.endswith(".xlsx"):
            input_file = os.path.join(input_folder, file_name)
            output_file = os.path.join(output_folder, os.path.splitext(file_name)[0] + ".pdf")
            excel_to_pdf(input_file, output_file)

if __name__ == "__main__":
    input_folder = "D:/123/456"
    output_folder = "D:/123_PDF"

    batch_convert_excel_to_pdf(input_folder, output_folder)
