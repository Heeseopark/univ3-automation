from runtime_paths import (
    get_base_dir,
    get_dir,
    get_downloads_dir,
    get_kakao_downloads_dir,
    get_env_file_path,
    get_google_credentials_candidates,
    get_google_sheet_id,
    get_screenshot_dirs,
)

import openpyxl
from openpyxl import load_workbook
from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import TextBlock, CellRichText
import re
from datetime import datetime, timedelta
import os
import sys
import win32com.client as win32

# 색상 코드 정의
try:
    import colorama
    colorama.init()
    MAGENTA = '\033[95m'
    CYAN = '\033[96m'
    WHITE = '\033[97m'
    RESET = '\033[0m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
except ImportError:
    MAGENTA = CYAN = WHITE = RESET = GREEN = YELLOW = RED = ''

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def find_sunday_month():
    """
    현재 날짜가 속한 주의 일요일이 속한 월을 찾는다.
    월요일~일요일 기준으로, 현재 날짜가 속한 주의 일요일의 월을 반환
    """
    today = datetime.now()

    # 이번 주 일요일 찾기 (월요일이 0, 일요일이 6)
    days_until_sunday = 6 - today.weekday()
    if today.weekday() == 6:  # 오늘이 일요일인 경우
        this_sunday = today
    else:
        this_sunday = today + timedelta(days=days_until_sunday)

    return this_sunday.month

def select_month():
    """
    사용자가 이번 달 또는 다음 달을 선택할 수 있는 메뉴를 제공
    """
    current_month = datetime.now().month
    current_year = datetime.now().year

    # 다음 달 계산
    if current_month == 12:
        next_month = 1
        next_year = current_year + 1
    else:
        next_month = current_month + 1
        next_year = current_year

    # 이번 주 일요일의 월 확인 (기본값)
    sunday_month = find_sunday_month()

    clear_screen()
    print("=" * 60)
    print(f"{CYAN}         큐티 월 선택{RESET}")
    print("=" * 60)
    print()

    # 선택 옵션 표시
    print(f"현재 날짜: {datetime.now().strftime('%Y년 %m월 %d일')}")
    print(f"이번 주 일요일의 월: {sunday_month}월\n")

    print("어느 달로 설정하시겠습니까?\n")

    options = [
        (current_month, f"{current_year}년 {current_month}월 (이번 달)"),
        (next_month, f"{next_year if next_month == 1 else current_year}년 {next_month}월 (다음 달)")
    ]

    # Windows 환경인지 확인
    if os.name == 'nt':
        try:
            import msvcrt

            selected_index = 0 if sunday_month == current_month else 1  # 기본값: 일요일 기준 월

            while True:
                # 메뉴 다시 표시
                clear_screen()
                print("=" * 60)
                print(f"{CYAN}         큐티 월 선택{RESET}")
                print("=" * 60)
                print()
                print(f"현재 날짜: {datetime.now().strftime('%Y년 %m월 %d일')}")
                print(f"이번 주 일요일의 월: {sunday_month}월 {YELLOW}(추천){RESET}\n")
                print("어느 달로 설정하시겠습니까?\n")

                for i, (month, desc) in enumerate(options):
                    if i == selected_index:
                        print(f"{CYAN}  ▶ {desc}{RESET}")
                    else:
                        print(f"    {desc}")

                print("\n" + "=" * 60)
                print("↑/↓: 선택 변경 | Enter: 확인 | ESC: 취소")

                key = msvcrt.getch()

                if key in [b'\x00', b'\xe0']:  # 특수 키
                    key = msvcrt.getch()
                    if key == b'H':  # 위 화살표
                        selected_index = max(0, selected_index - 1)
                    elif key == b'P':  # 아래 화살표
                        selected_index = min(len(options) - 1, selected_index + 1)
                elif key == b'\r':  # Enter
                    selected_month = options[selected_index][0]
                    print(f"\n{GREEN}✓ {selected_month}월이 선택되었습니다.{RESET}")
                    return selected_month
                elif key == b'\x1b':  # ESC
                    print(f"\n{YELLOW}취소되었습니다. 기본값({sunday_month}월)을 사용합니다.{RESET}")
                    return sunday_month

        except ImportError:
            # msvcrt가 없으면 일반 입력 방식 사용
            pass

    # Unix/Linux/Mac 또는 msvcrt가 없는 경우
    for i, (month, desc) in enumerate(options):
        print(f"  {i+1}. {desc}")

    print(f"\n기본값: {sunday_month}월 (Enter를 누르면 기본값 선택)")
    print("=" * 60)

    while True:
        choice = input("선택 (1 또는 2, Enter=기본값): ").strip()

        if choice == "":
            print(f"\n{GREEN}✓ 기본값 {sunday_month}월이 선택되었습니다.{RESET}")
            return sunday_month
        elif choice == "1":
            print(f"\n{GREEN}✓ {current_month}월이 선택되었습니다.{RESET}")
            return current_month
        elif choice == "2":
            print(f"\n{GREEN}✓ {next_month}월이 선택되었습니다.{RESET}")
            return next_month
        else:
            print(f"{RED}잘못된 입력입니다. 1 또는 2를 입력하세요.{RESET}")

def process_excel_file(file_path, selected_month=None):
    """
    엑셀 파일을 처리하여 큐티 시트의 H1 셀 날짜를 수정하고 저장한다.
    """
    try:
        # 월이 지정되지 않은 경우 사용자에게 선택하게 함
        if selected_month is None:
            selected_month = select_month()

        # COM을 사용하여 서식 유지하며 수정
        return update_with_com(file_path, selected_month)

    except Exception as e:
        print(f"엑셀 파일 처리 중 오류 발생: {e}")
        return False

def update_with_com(file_path, new_month):
    """
    COM을 사용하여 서식을 유지하며 H1 셀의 월 정보를 수정
    """
    try:
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False

        # 파일 열기
        abs_path = os.path.abspath(file_path)
        wb = excel.Workbooks.Open(abs_path)

        # 큐티 시트 찾기 (첫 번째 시트가 큐티 시트라고 가정)
        quti_sheet = None
        for sheet in wb.Sheets:
            if '큐티' in sheet.Name:
                quti_sheet = sheet
                print(f"{GREEN}큐티 시트를 찾았습니다: {sheet.Name}{RESET}")
                break

        # 큐티 시트가 없으면 첫 번째 시트 사용
        if not quti_sheet:
            quti_sheet = wb.Sheets[1]
            print(f"{YELLOW}첫 번째 시트를 사용합니다: {quti_sheet.Name}{RESET}")

        # H1 셀 처리
        cell = quti_sheet.Range('H1')
        text = cell.Value

        if text:
            print(f"현재 H1 셀 값: {text}")

            # 월 정보 찾기
            import re
            month_pattern = r'(\d{1,2})월'
            match = re.search(month_pattern, str(text))

            if match:
                old_month = match.group(1)

                # H1 셀 전체를 새로운 월로 교체 (서식 유지)
                # Characters를 사용하여 텍스트만 변경
                new_text = f"{new_month}월"

                # 기존 서식 정보 저장
                font_name = cell.Font.Name
                font_size = cell.Font.Size
                font_bold = cell.Font.Bold
                font_color = cell.Font.Color
                font_italic = cell.Font.Italic
                font_underline = cell.Font.Underline

                # 텍스트 변경
                cell.Value = new_text

                # 서식 재적용
                cell.Font.Name = font_name
                cell.Font.Size = font_size
                cell.Font.Bold = font_bold
                cell.Font.Color = font_color
                cell.Font.Italic = font_italic
                cell.Font.Underline = font_underline

                print(f"{old_month}월 -> {new_month}월로 변경 (서식 유지)")
            else:
                # 월 정보가 없으면 그냥 새로운 월로 설정
                print(f"H1 셀에 월 정보가 없어서 {new_month}월로 설정합니다.")
                cell.Value = f"{new_month}월"
        else:
            # H1 셀이 비어있으면 새로운 월 설정
            print(f"H1 셀이 비어있어서 {new_month}월로 설정합니다.")
            cell.Value = f"{new_month}월"

        # 파일 저장
        wb.Save()
        print(f"엑셀 파일이 저장되었습니다: {os.path.basename(file_path)}")

        wb.Close()
        excel.Quit()

        return new_month

    except Exception as e:
        print(f"COM 처리 중 오류: {e}")
        try:
            wb.Close()
            excel.Quit()
        except:
            pass
        return False

def export_to_pdf(excel_file_path, month=None):
    """
    엑셀 파일을 PDF로 변환한다. (Windows 전용, Excel 설치 필요)
    """
    try:
        # PDF 파일 경로 설정 (같은 폴더에 저장, 월 정보 포함)
        if month:
            base_dir = os.path.dirname(excel_file_path)
            pdf_filename = f"{month}월 큐티 구매.pdf"
            pdf_file_path = os.path.join(base_dir, pdf_filename)
        else:
            pdf_file_path = excel_file_path.replace('.xlsx', '.pdf').replace('.xls', '.pdf')

        # Excel 응용 프로그램 시작
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False

        # 절대 경로로 변환
        abs_excel_path = os.path.abspath(excel_file_path)
        abs_pdf_path = os.path.abspath(pdf_file_path)

        # 엑셀 파일 열기
        wb = excel.Workbooks.Open(abs_excel_path)

        # 큐티 시트 찾기
        quti_sheet = None
        for sheet in wb.Sheets:
            if '큐티' in sheet.Name:
                quti_sheet = sheet
                break

        # 큐티 시트가 없으면 첫 번째 시트 사용
        if not quti_sheet:
            quti_sheet = wb.Sheets[1]

        if quti_sheet:
            # 큐티 시트만 PDF로 내보내기
            quti_sheet.ExportAsFixedFormat(0, abs_pdf_path)
            print(f"큐티 시트가 PDF로 저장되었습니다: {os.path.basename(pdf_file_path)}")
        else:
            # 전체 워크북을 PDF로 내보내기
            wb.ExportAsFixedFormat(0, abs_pdf_path)
            print(f"전체 엑셀 파일이 PDF로 저장되었습니다: {os.path.basename(pdf_file_path)}")

        # 파일 닫기
        wb.Close(False)
        excel.Quit()

        return True

    except Exception as e:
        print(f"PDF 변환 중 오류 발생: {e}")
        print("Windows에서 Excel이 설치되어 있어야 PDF 변환이 가능합니다.")
        return False

def main():
    """메인 실행 함수"""
    # 프린트 폴더에서 '교재 큐티 구매.xlsx' 파일 찾기
    base_dir = get_base_dir()
    target_file = os.path.join(base_dir, "프린트", "교재 큐티 구매.xlsx")

    if not os.path.exists(target_file):
        print(f"{RED}파일을 찾을 수 없습니다: 교재 큐티 구매.xlsx{RESET}")
        print(f"경로: {target_file}")
        return False

    # 사용자에게 월 선택 받기
    selected_month = select_month()

    print(f"\n{CYAN}찾은 엑셀 파일: 교재 큐티 구매.xlsx{RESET}")
    print(f"처리 중: {selected_month}월로 업데이트...")

    # 엑셀 파일 처리 (선택한 월 전달)
    result = process_excel_file(target_file, selected_month)
    if result:
        # PDF로 변환 (월 정보 전달)
        pdf_success = export_to_pdf(target_file, month=result)
        if pdf_success:
            print(f"\n{GREEN}✓ 모든 작업이 성공적으로 완료되었습니다!{RESET}")
            print(f"  - 엑셀 파일: {selected_month}월로 업데이트 완료")
            print(f"  - PDF 파일: {selected_month}월 큐티 구매.pdf 생성 완료")
            return True
        else:
            print(f"\n{RED}PDF 변환에 실패했습니다.{RESET}")
            return False
    else:
        print(f"{YELLOW}교재 큐티 구매.xlsx 파일 처리를 건너뜁니다.{RESET}")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
