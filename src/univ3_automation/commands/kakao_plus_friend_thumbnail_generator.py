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

import os
import sys
from datetime import datetime, timedelta
from pptx import Presentation
from pptx.util import Pt
from PIL import Image
import win32com.client
import pythoncom
import pytz

# ====================================
# 설정 - 수정할 슬라이드 번호 (0부터 시작)
# ====================================
SLIDE_NUMBER = 0

# 색상 코드 정의
try:
    import colorama
    colorama.init()
    MAGENTA = '\033[95m'
    CYAN = '\033[96m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    RESET = '\033[0m'
except ImportError:
    MAGENTA = CYAN = GREEN = YELLOW = RED = RESET = ''

def get_sunday_of_current_week():
    """현재 한국 시간 기준으로 이번 주 일요일 날짜를 반환"""
    # 한국 시간대 설정
    kst = pytz.timezone('Asia/Seoul')
    now_kst = datetime.now(kst)

    # 현재 날짜의 요일 (0=월요일, 6=일요일)
    weekday = now_kst.weekday()

    # 이번 주 일요일 계산
    # 월요일(0)부터 일요일까지의 차이 계산
    if weekday == 6:  # 이미 일요일인 경우
        sunday = now_kst
    else:
        # 다음 일요일까지 남은 일수
        days_until_sunday = 6 - weekday
        sunday = now_kst + timedelta(days=days_until_sunday)

    return sunday

def modify_ppt_date(meeting_name="주일집회"):
    """PPT 파일의 날짜와 집회명을 이번 주 일요일로 수정하고 슬라이드를 이미지로 추출"""

    print("====================================")
    print("카톡플친 PPT 처리 스크립트")
    print(f"수정할 슬라이드 번호: {SLIDE_NUMBER}")
    print(f"집회 이름: {MAGENTA}{meeting_name}{RESET}")
    print("====================================\n")

    # 카톡플친 디렉토리 경로 설정
    parent_dir = get_base_dir()
    kakao_dir = os.path.join(parent_dir, "카톡플친")
    ppt_file_path = os.path.join(kakao_dir, "카톡플친 커버.pptx")

    if not os.path.exists(ppt_file_path):
        print(f"오류: {ppt_file_path} 파일을 찾을 수 없습니다.")
        return False

    try:
        # PPT 파일 열기
        print(f"PPT 파일을 열고 있습니다: {ppt_file_path}")
        try:
            prs = Presentation(ppt_file_path)
        except PermissionError:
            print("\n오류: PPT 파일에 접근할 수 없습니다.")
            print("다음을 확인해주세요:")
            print("1. PowerPoint에서 파일이 열려있다면 닫아주세요.")
            print("2. 다른 프로그램에서 파일을 사용 중이라면 종료해주세요.")
            print("3. 파일이 읽기 전용이 아닌지 확인해주세요.")
            return False

        # 이번 주 일요일 날짜 가져오기
        sunday = get_sunday_of_current_week()
        date_str = sunday.strftime("%m%d")  # MMDD 형식 (예: 0922)
        date_display = f"{sunday.month}월 {sunday.day}일"  # 표시용 (예: 9월 22일)

        print(f"이번 주 일요일: {date_display}")
        print(f"파일명에 사용될 날짜: {date_str}\n")

        # 지정된 슬라이드의 날짜 텍스트 수정
        slide_idx = SLIDE_NUMBER - 1  # 0-based index

        if slide_idx >= len(prs.slides):
            print(f"오류: 슬라이드 {SLIDE_NUMBER}번을 찾을 수 없습니다. 전체 슬라이드 수: {len(prs.slides)}")
            return False

        slide = prs.slides[slide_idx]
        print(f"슬라이드 {SLIDE_NUMBER}번의 날짜를 수정합니다...")

        # 슬라이드의 모든 텍스트 프레임에서 날짜 패턴 찾기
        date_modified = False
        import re
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls

        for shape in slide.shapes:
            if hasattr(shape, "text_frame"):
                text_frame = shape.text_frame
                # 전체 텍스트 프레임 내용 가져오기
                full_text = text_frame.text
                if "년" in full_text and "월" in full_text and "일" in full_text:
                    print(f"  기존 텍스트 프레임:\n{full_text}")

                    # 각 paragraph를 처리
                    for paragraph in text_frame.paragraphs:
                        para_text = paragraph.text
                        if "년" in para_text and "월" in para_text and "일" in para_text:
                            # 날짜 부분만 찾아서 교체
                            pattern = r"\d{4}년\s+\d{1,2}월\s+\d{1,2}일"
                            match = re.search(pattern, para_text)

                            if match:
                                # 날짜 위치 찾기
                                start_pos = match.start()
                                end_pos = match.end()
                                replacement = f"{sunday.year}년 {sunday.month}월 {sunday.day}일"

                                # Run 레벨에서 직접 텍스트 수정
                                if paragraph.runs:
                                    # 모든 run을 순회하면서 위치 계산
                                    current_pos = 0
                                    for run in paragraph.runs:
                                        run_text = run.text
                                        run_len = len(run_text)
                                        run_start = current_pos
                                        run_end = current_pos + run_len

                                        # 이 run이 날짜를 포함하는지 확인
                                        if run_start < end_pos and run_end > start_pos:
                                            # 날짜가 이 run에 포함됨
                                            if run_start <= start_pos and run_end >= end_pos:
                                                # 날짜가 완전히 이 run 안에 있음
                                                before = run_text[:start_pos - run_start]
                                                after = run_text[end_pos - run_start:]
                                                run.text = before + replacement + after
                                            elif run_start <= start_pos:
                                                # 날짜 시작이 이 run에 있음
                                                before = run_text[:start_pos - run_start]
                                                run.text = before + replacement
                                            elif run_end >= end_pos:
                                                # 날짜 끝이 이 run에 있음
                                                after = run_text[end_pos - run_start:]
                                                run.text = after
                                            else:
                                                # 이 run은 날짜 중간에 있음 - 비움
                                                run.text = ""

                                        current_pos += run_len
                                else:
                                    # run이 없는 경우 paragraph 레벨에서 텍스트 설정
                                    new_para_text = para_text[:start_pos] + replacement + para_text[end_pos:]
                                    paragraph.text = new_para_text

                                print(f"  변경된 날짜: {replacement}")
                                date_modified = True

        if not date_modified:
            print(f"  경고: 슬라이드 {SLIDE_NUMBER}번에서 날짜 텍스트를 찾을 수 없습니다.")

        # 집회명 수정 (날짜가 있는 텍스트 박스에서만 수정)
        meeting_modified = False
        print(f"\n슬라이드 {SLIDE_NUMBER}번의 집회명을 수정합니다...")

        for shape in slide.shapes:
            if hasattr(shape, "text_frame"):
                text_frame = shape.text_frame
                full_text = text_frame.text
                # 날짜가 포함된 텍스트 박스에서만 집회명 수정
                if "년" in full_text and "월" in full_text and "일" in full_text:
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            if "집회" in run.text:
                                # 집회가 포함된 텍스트 찾기 (예: 주일집회, 개강사경회 등)
                                old_text = run.text
                                # "XX집회" 패턴을 찾아서 meeting_name으로 교체
                                new_text = re.sub(r'\S*집회', meeting_name, old_text)
                                if old_text != new_text:
                                    run.text = new_text
                                    print(f"  집회명 변경: {old_text} -> {new_text}")
                                    meeting_modified = True

        if not meeting_modified:
            print(f"  경고: 슬라이드 {SLIDE_NUMBER}번에서 집회명 텍스트를 찾을 수 없습니다.")

        # 기존 PPT 파일에 덮어쓰기로 저장
        try:
            prs.save(ppt_file_path)
            print(f"\nPPT 파일 저장 완료: {ppt_file_path}")
        except PermissionError:
            print("\n오류: PPT 파일을 저장할 수 없습니다.")
            print("PowerPoint에서 파일이 열려있다면 닫고 다시 시도해주세요.")
            return False

        # PowerPoint COM 객체를 사용하여 슬라이드를 이미지로 추출
        print(f"슬라이드 {SLIDE_NUMBER}번을 이미지로 추출합니다...")

        pythoncom.CoInitialize()

        try:
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            powerpoint.Visible = 1  # PowerPoint는 반드시 Visible이어야 함
            powerpoint.WindowState = 2  # ppWindowMinimized - 최소화 상태로 실행

            # 절대 경로 사용
            presentation = powerpoint.Presentations.Open(ppt_file_path)

            # 슬라이드를 이미지로 저장
            output_image = f"카톡플친 커버 {date_str}.png"
            output_image_path = os.path.join(kakao_dir, output_image)

            # 슬라이드를 원본 크기로 내보내기

            # 슬라이드 내보내기 (원본 크기)
            presentation.Slides[SLIDE_NUMBER].Export(output_image_path, "PNG", 2268, 1134)

            print(f"이미지 저장 완료: {output_image_path}")

            # PowerPoint 종료
            presentation.Close()
            powerpoint.Quit()

        finally:
            pythoncom.CoUninitialize()

        print("\n작업이 성공적으로 완료되었습니다!")
        return True

    except Exception as e:
        print(f"\n오류 발생: {str(e)}")
        return False

if __name__ == "__main__":
    # 집회 이름 입력
    print(f"{CYAN}집회 이름을 입력하세요 (예: 개강사경회){RESET}")
    print(f"{YELLOW}빈 값 입력 시 '주일집회'로 설정됩니다.{RESET}")
    meeting_name_input = input("집회 이름: ").strip()
    meeting_name = meeting_name_input if meeting_name_input else "주일집회"

    # 실행
    success = modify_ppt_date(meeting_name)
    sys.exit(0 if success else 1)
