import pandas as pd
from bs4 import BeautifulSoup
import os

# 파일 경로 설정 (사용자 환경에 맞게 수정하세요)
INPUT_HTML_FILE = 'cve.html'
OUTPUT_EXCEL_FILE = 'cve_report_output.xlsx'

def html_to_excel_with_format(html_file_path, excel_file_path):
    """
    HTML 파일을 읽어 엑셀로 변환하며, 줄바꿈과 하이퍼링크를 유지합니다.
    """
    
    if not os.path.exists(html_file_path):
        print(f"오류: {html_file_path} 파일을 찾을 수 없습니다.")
        return

    print("1. HTML 파일 파싱 중...")
    with open(html_file_path, 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f, 'html.parser')

    # HTML 내의 첫 번째 테이블 찾기
    table = soup.find('table')
    if not table:
        print("오류: HTML 파일 내에서 <table> 태그를 찾을 수 없습니다.")
        return

    # --- 데이터 추출 로직 ---
    data = []
    
    # 헤더(제목) 추출
    headers = []
    thead = table.find('thead')
    if thead:
        headers = [th.get_text(strip=True) for th in thead.find_all('th')]
    else:
        # thead가 없으면 첫 번째 tr을 헤더로 가정
        first_row = table.find('tr')
        if first_row:
            headers = [th.get_text(strip=True) for th in first_row.find_all(['th', 'td'])]

    # 행(Row) 데이터 추출
    rows = table.find_all('tr')
    # 헤더가 별도로 있었거나 첫 줄을 헤더로 썼다면 rows에서 제외 필요 (상황에 따라 조정)
    if thead:
        rows_to_process = table.find('tbody').find_all('tr') if table.find('tbody') else rows
    else:
        rows_to_process = rows[1:] # 첫 줄 제외

    print(f"2. {len(rows_to_process)}개의 행 데이터 변환 중...")

    for tr in rows_to_process:
        row_data = []
        cells = tr.find_all(['td', 'th'])
        
        for cell in cells:
            # 1. 줄바꿈 처리: <br> 태그를 실제 줄바꿈 문자(\n)로 변환
            # get_text(separator="\n")를 사용하면 <br>과 <p>등을 줄바꿈으로 바꿔줍니다.
            cell_text = cell.get_text(separator="\n", strip=True)
            
            # 2. 하이퍼링크 처리
            # 셀 안에 <a> 태그가 있는지 확인
            link = cell.find('a')
            if link and link.get('href'):
                url = link.get('href')
                # 엑셀의 HYPERLINK 함수 형식으로 변환: =HYPERLINK("주소", "표시텍스트")
                # 주의: 엑셀 수식은 255자를 넘으면 잘릴 수 있으나, 최신 엑셀은 대부분 지원합니다.
                # 데이터가 너무 복잡하면 그냥 텍스트 + URL로 병기하는 것이 안전할 수 있습니다.
                
                # 여기서는 'Excel Formula' 방식 대신, 나중에 xlsxwriter로 쓸 때 URL을 인식시키는 방식을 위해
                # 튜플 형태로 (텍스트, URL)을 저장하거나, 
                # 가장 확실한 방법인 'Excel 수식 문자열'을 만듭니다.
                
                # 수식 내 따옴표(") 이스케이프 처리
                safe_url = url.replace('"', '""')
                safe_text = cell_text.replace('"', '""')
                
                # 텍스트가 너무 길면 수식 오류가 날 수 있으므로, 
                # URL이 있는 경우엔 [텍스트](URL) 형태로 저장하거나 수식을 씁니다.
                # 여기서는 가장 깔끔한 수식 적용을 시도합니다.
                cell_val = f'=HYPERLINK("{safe_url}", "{safe_text}")'
            else:
                cell_val = cell_text
            
            row_data.append(cell_val)
        
        # 행 길이가 헤더와 다를 경우 빈 값 채우기 (오류 방지)
        if len(row_data) < len(headers):
            row_data += [''] * (len(headers) - len(row_data))
            
        data.append(row_data)

    # DataFrame 생성
    df = pd.DataFrame(data, columns=headers)

    # --- 엑셀 저장 및 서식 적용 (XlsxWriter 엔진 사용) ---
    print("3. 엑셀 파일 생성 및 서식 적용 중...")
    
    # Pandas의 ExcelWriter를 사용하여 서식 제어
    with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='CVE Report')
        
        workbook = writer.book
        worksheet = writer.sheets['CVE Report']

        # 서식 정의
        # 1. 줄바꿈(Wrap Text) 서식
        wrap_format = workbook.add_format({
            'text_wrap': True,
            'valign': 'top',
            'border': 1
        })
        
        # 2. 헤더 서식
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })

        # 데이터가 있는 전체 영역에 줄바꿈 서식 적용
        # (Pandas가 데이터를 쓴 후, 서식을 덮어씌웁니다)
        for row_idx, row in enumerate(data):
            for col_idx, value in enumerate(row):
                # 엑셀은 0-based index, 헤더가 있으므로 row_idx + 1
                worksheet.write(row_idx + 1, col_idx, value, wrap_format)

        # 열 너비 자동 조정 (대략적으로)
        # Description 같은 긴 컬럼은 넓게, 나머지는 적당히
        for i, col in enumerate(headers):
            if 'description' in col.lower() or 'summary' in col.lower() or '내용' in col:
                worksheet.set_column(i, i, 50) # 너비 50
            elif 'id' in col.lower() or 'cve' in col.lower():
                worksheet.set_column(i, i, 15) # 너비 15
            else:
                worksheet.set_column(i, i, 20) # 기본 너비

    print(f"완료! 파일이 저장되었습니다: {excel_file_path}")

# 실행
if __name__ == "__main__":
    # 라이브러리 설치 필요: pip install pandas beautifulsoup4 xlsxwriter openpyxl
    html_to_excel_with_format(INPUT_HTML_FILE, OUTPUT_EXCEL_FILE)
