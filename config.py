import os

# 현재 스크립트 파일 경로
project_path = os.path.dirname(os.path.abspath(__file__))
doc_path = f"{project_path}/doc"
generated_doc_path = f"{project_path}/generated_doc"

# 파일명
base_info_file_name = "0000_base_info_file.xlsx"
test_excel_file = "0100_test_excel_file.xlsx"
test_docx_file = "0500_test_docx_file.docx"