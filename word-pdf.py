import os
from win32com import client

def word_to_pdf(input_file, output_file):
    word = client.Dispatch("Word.Application")
    doc = word.Documents.Open(input_file)
    doc.SaveAs(output_file, FileFormat=17)  # 17 represents PDF format
    doc.Close()
    word.Quit()

def batch_convert_word_to_pdf(input_folder, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for file_name in os.listdir(input_folder):
        if file_name.endswith(".docx") or file_name.endswith(".doc"):
            input_file = os.path.join(input_folder, file_name)
            output_file = os.path.join(output_folder, os.path.splitext(file_name)[0] + ".pdf")
            word_to_pdf(input_file, output_file)

if __name__ == "__main__":
    input_folder = "D:/123/222"
    output_folder = "D:/123/333"

    batch_convert_word_to_pdf(input_folder, output_folder)
