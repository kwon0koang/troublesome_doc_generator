import os
import openpyxl
from docx import Document

# 현재 스크립트 파일 경로 가져오기
project_path = os.path.dirname(os.path.abspath(__file__))
doc_path = f"{project_path}/doc"
generated_doc_path = f"{project_path}/generated_doc"

# 파일명
base_info_file_name = "0000_base_info_file.xlsx"
test_excel_file = "0100_test_excel_file.xlsx"
test_docx_file = "0500_test_docx_file.docx"
 
# ====================================================================================================================================================================================================================================================================================

class DocInfo:
    def __init__(self, sr_title, sr_no, developer, author, approver, complete_dev_date, deploy_date, developer_test_infos):
        self.sr_title = sr_title
        self.sr_no = sr_no
        self.developer = developer
        self.author = author
        self.approver = approver
        self.complete_dev_date = complete_dev_date
        self.deploy_date = deploy_date
        self.developer_test_infos = developer_test_infos
        
class DeveloperTestInfo:
    def __init__(self, test_content, developer, approver):
        self.test_content = test_content
        self.developer = developer
        self.approver = approver
 
# ====================================================================================================================================================================================================================================================================================

def create_directory(directory_path):
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)
        print(f"'{directory_path}' 생성")
    else:
        print(f"'{directory_path}' 이미 존재")

def get_doc_info() -> DocInfo:
    # 엑셀 파일 읽기
    excel_path = f"{doc_path}/{base_info_file_name}"
    wb = openpyxl.load_workbook(excel_path)
    sheet = wb.active

    # 값 가져오기
    sr_title = sheet['B2'].value
    sr_no = sheet['B3'].value
    developer = sheet['B8'].value
    author = sheet['B9'].value
    approver = sheet['B10'].value
    complete_dev_date = sheet['E2'].value
    deploy_date = sheet['E3'].value
    developer_test_infos = get_developer_test_infos(sheet, approver)
    doc_info = DocInfo(sr_title, sr_no, developer, author, approver, complete_dev_date, deploy_date, developer_test_infos)
    
    # 파일 닫기
    wb.close()
    
    return doc_info

def get_developer_test_infos(sheet, approver) -> list[DeveloperTestInfo]:
    try:
        # g열과 h열의 데이터 읽기
        g_column = sheet['G']
        h_column = sheet['H']

        # 데이터를 저장할 리스트 초기화
        data_list = []

        # 열의 길이 중 작은 길이까지만 반복
        # for g_cell, h_cell in zip(g_column, h_column):
        for g_cell, h_cell in zip(g_column[1:], h_column[1:]): # 2행부터 테스트 데이터들 존재. 2행부터 데이터 읽기
            # 둘 중 하나라도 값이 비어 있으면 반복 종료
            if g_cell.value is None or h_cell.value is None:
                break

            # DeveloperTestInfo 객체 생성 후 데이터 리스트에 추가
            data_list.append(DeveloperTestInfo(test_content=g_cell.value, developer=h_cell.value, approver=approver))

        return data_list

    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

def generate_b(doc_info: DocInfo):
    # 엑셀 파일 읽기
    excel_path = f"{doc_path}/{test_excel_file}"
    wb = openpyxl.load_workbook(excel_path)
    b_sheet = wb['BBB']  # BBB 탭 선택

    # 값 채우기
    b_sheet['A1'] = doc_info.sr_title
    b_sheet['A2'] = doc_info.sr_no
    b_sheet['A4'] = doc_info.developer
    b_sheet['B6'] = doc_info.author
    b_sheet['C10'] = doc_info.approver
    
    # 데이터를 D5셀부터 채우기
    c_sheet = wb['CCC']  # CCC 탭 선택
    for row_index, test_info in enumerate(doc_info.developer_test_infos, start=4):  # 엑셀은 1부터 시작하지만, 리스트는 0부터 시작하므로 4부터 시작
        c_sheet.cell(row=row_index, column=4, value=test_info.test_content)
        c_sheet.cell(row=row_index, column=5, value=test_info.developer)
        c_sheet.cell(row=row_index, column=6, value=test_info.approver)
        c_sheet.cell(row=row_index, column=7, value=doc_info.complete_dev_date)
    
    # 파일 저장
    wb.save(f"{generated_doc_path}/{test_excel_file}")
    
    # 파일 닫기
    wb.close()
    
def generate_c(doc_info: DocInfo):
    # docx 파일 읽기
    docx_path = f"{doc_path}/{test_docx_file}"
    docx = Document(docx_path)

    # 파일에서 값 찾아 치환
    for paragraph in docx.paragraphs:
        if "AAA" in paragraph.text:
            paragraph.text = paragraph.text.replace(paragraph.text, doc_info.developer)
        if "BBB" in paragraph.text:
            paragraph.text = paragraph.text.replace(paragraph.text, doc_info.author)
        if "CCC" in paragraph.text:
            paragraph.text = paragraph.text.replace(paragraph.text, doc_info.approver)

    # 파일 저장
    docx.save(f"{generated_doc_path}/{test_docx_file}")
    
# ====================================================================================================================================================================================================================================================================================

if __name__ == "__main__":
    # 생성된 파일 저장할 폴더 생성
    create_directory(generated_doc_path)

    # 기본 정보 가져오기
    doc_info = get_doc_info()

    # 엑셀 파일 생성
    generate_b(doc_info)

    # docx 파일 생성
    generate_c(doc_info)