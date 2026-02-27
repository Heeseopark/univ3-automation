#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
행정광고 PPT에 빨간색 직사각형 텍스트 박스 추가 스크립트
돌아오는 주일 행정광고 PPT에 투명도 50% 빨간색 직사각형과 흰색 텍스트를 추가합니다.
"""

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
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, PP_PARAGRAPH_ALIGNMENT
from pptx.dml.color import RGBColor
import pytz
import re
import win32com.client
import pythoncom

# Windows 콘솔 UTF-8 설정
if sys.platform == 'win32':
    # Windows 콘솔 코드 페이지를 UTF-8로 설정
    os.system('chcp 65001 >nul 2>&1')

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


def get_this_sunday():
    """현재 한국 시간 기준으로 이번 주 일요일 날짜를 반환"""
    kst = pytz.timezone('Asia/Seoul')
    now_kst = datetime.now(kst)
    weekday = now_kst.weekday()

    if weekday == 6:  # 이미 일요일인 경우
        sunday = now_kst
    else:
        days_until_sunday = 6 - weekday
        sunday = now_kst + timedelta(days=days_until_sunday)

    return sunday


def find_admin_ppt(admin_dir, sunday_str):
    """이번 주 행정광고 PPT 파일 찾기"""
    ppt_name = f"{sunday_str} 행정광고.pptx"
    ppt_path = os.path.join(admin_dir, ppt_name)

    if os.path.exists(ppt_path):
        return ppt_path
    return None


def open_ppt_in_powerpoint(ppt_path):
    """PowerPoint로 PPT 파일 열기"""
    print(f"\n{CYAN}PowerPoint로 PPT 파일을 여는 중...{RESET}")
    pythoncom.CoInitialize()

    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1  # PowerPoint 창 표시
        powerpoint.WindowState = 1  # ppWindowNormal - 일반 크기로 표시

        # 절대 경로로 변환
        abs_ppt_path = os.path.abspath(ppt_path)

        # 이미 열린 파일인지 확인
        for presentation in powerpoint.Presentations:
            if os.path.abspath(presentation.FullName) == abs_ppt_path:
                print(f"{YELLOW}PPT 파일이 이미 열려있습니다.{RESET}")
                return powerpoint, presentation

        # 파일 열기
        presentation = powerpoint.Presentations.Open(abs_ppt_path)
        print(f"{GREEN}PPT 파일이 열렸습니다. PowerPoint 창에서 슬라이드를 확인하세요.{RESET}")
        return powerpoint, presentation

    except Exception as e:
        print(f"{RED}PPT 파일 열기 실패: {str(e)}{RESET}")
        pythoncom.CoUninitialize()
        return None, None


def close_ppt_in_powerpoint(powerpoint, presentation, save=True):
    """PowerPoint 프레젠테이션 닫기 (PowerPoint는 닫지 않음)"""
    try:
        if presentation and save:
            presentation.Save()
            print(f"\n{GREEN}PPT 파일이 저장되었습니다.{RESET}")
        # presentation.Close()는 호출하지 않음 - 사용자가 계속 볼 수 있도록
    except Exception as e:
        print(f"{YELLOW}저장 중 오류: {str(e)}{RESET}")
    finally:
        pythoncom.CoUninitialize()


def get_position_input():
    """직사각형 위치 입력받기"""
    print(f"\n{CYAN}직사각형 위치를 선택하세요:{RESET}")
    print("  1. 좌상 (왼쪽 위)")
    print("  2. 우상 (오른쪽 위)")
    print("  3. 좌하 (왼쪽 아래)")
    print("  4. 우하 (오른쪽 아래)")

    while True:
        try:
            choice = input("위치 번호 (1-4): ").strip()
            if choice in ['1', '2', '3', '4']:
                return int(choice)
            else:
                print(f"{RED}1-4 사이의 숫자를 입력해주세요.{RESET}")
        except ValueError:
            print(f"{RED}올바른 숫자를 입력해주세요.{RESET}")


def calculate_rectangle_position(slide_width, slide_height, position_choice):
    """
    위치 선택에 따라 직사각형의 좌표 계산

    Args:
        slide_width: 슬라이드 너비 (EMUs)
        slide_height: 슬라이드 높이 (EMUs)
        position_choice: 위치 선택 (1: 좌상, 2: 우상, 3: 좌하, 4: 우하)

    Returns:
        (left, top, width, height) 튜플 (EMUs)
    """
    # 직사각형 크기 설정 (슬라이드의 약 40%)
    rect_width = slide_width * 0.4
    rect_height = slide_height * 0.4

    if position_choice == 1:  # 좌상
        left = slide_width * 0.05
        top = slide_height * 0.05
    elif position_choice == 2:  # 우상
        left = slide_width - rect_width - slide_width * 0.05
        top = slide_height * 0.05
    elif position_choice == 3:  # 좌하
        left = slide_width * 0.05
        top = slide_height - rect_height - slide_height * 0.05
    elif position_choice == 4:  # 우하
        left = slide_width - rect_width - slide_width * 0.05
        top = slide_height - rect_height - slide_height * 0.05
    else:
        # 기본값: 우하
        left = slide_width - rect_width - slide_width * 0.05
        top = slide_height - rect_height - slide_height * 0.05

    return int(left), int(top), int(rect_width), int(rect_height)


def add_text_box_to_slide_com(presentation, slide_number, position_choice, text_content):
    """
    win32com을 사용하여 슬라이드에 빨간색 직사각형과 흰색 텍스트 추가

    Args:
        presentation: win32com Presentation 객체
        slide_number: 슬라이드 번호 (1부터 시작)
        position_choice: 위치 선택 (1-4)
        text_content: 텍스트 내용

    Returns:
        성공 여부 (bool)
    """
    try:
        # 슬라이드 가져오기 (1부터 시작)
        if slide_number < 1 or slide_number > presentation.Slides.Count:
            print(f"{RED}오류: 슬라이드 번호가 범위를 벗어났습니다. (1-{presentation.Slides.Count}){RESET}")
            return False

        slide = presentation.Slides(slide_number)

        # 슬라이드 크기 가져오기 (포인트 단위)
        slide_width = presentation.PageSetup.SlideWidth
        slide_height = presentation.PageSetup.SlideHeight

        # 직사각형 위치 계산 (EMUs를 포인트로 변환)
        # python-pptx는 EMUs, win32com은 포인트 사용
        left_emu, top_emu, width_emu, height_emu = calculate_rectangle_position(
            slide_width, slide_height, position_choice
        )

        # 직사각형 도형 추가 (msoShapeRectangle = 1)
        shape = slide.Shapes.AddShape(
            1,  # msoShapeRectangle
            left_emu, top_emu, width_emu, height_emu
        )

        # 직사각형 스타일 설정
        # 빨간색 배경 (투명도 50%)
        shape.Fill.Solid()
        shape.Fill.ForeColor.RGB = 255  # 빨간색 (RGB(255, 0, 0))
        shape.Fill.Transparency = 0.5  # 50% 투명도

        # 윤곽선 없음
        shape.Line.Visible = 0  # msoFalse

        # 텍스트 추가
        text_frame = shape.TextFrame
        text_frame.WordWrap = -1  # msoTrue
        text_frame.MarginLeft = 10
        text_frame.MarginRight = 10
        text_frame.MarginTop = 10
        text_frame.MarginBottom = 10

        # 텍스트 내용 설정
        text_range = text_frame.TextRange
        text_range.Text = text_content

        # 텍스트 스타일 설정
        text_range.Font.Name = '맑은 고딕'
        text_range.Font.Size = 18
        text_range.Font.Bold = -1  # msoTrue
        text_range.Font.Color.RGB = 16777215  # 흰색 (RGB(255, 255, 255))

        # 텍스트 정렬
        text_range.ParagraphFormat.Alignment = 2  # ppAlignCenter
        text_frame.VerticalAnchor = 3  # msoAnchorMiddle

        # 변경사항 저장
        presentation.Save()

        print(f"{GREEN}슬라이드 {slide_number}에 텍스트 박스가 추가되었습니다.{RESET}")
        return True

    except Exception as e:
        print(f"{RED}오류 발생: {str(e)}{RESET}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """메인 함수"""
    print("=" * 60)
    print(f"{CYAN}행정광고 PPT 텍스트 박스 추가 스크립트{RESET}")
    print("=" * 60 + "\n")

    # 디렉토리 설정
    parent_dir = get_base_dir()
    admin_dir = os.path.join(parent_dir, "행정광고")

    if not os.path.exists(admin_dir):
        print(f"{RED}오류: {admin_dir} 디렉토리를 찾을 수 없습니다.{RESET}")
        return False

    # 이번주 일요일 날짜
    this_sunday = get_this_sunday()
    sunday_str = this_sunday.strftime("%m%d")

    print(f"돌아오는 주일: {this_sunday.strftime('%Y년 %m월 %d일')} ({sunday_str})")

    # PPT 파일 찾기
    ppt_path = find_admin_ppt(admin_dir, sunday_str)

    if not ppt_path:
        print(f"\n{RED}오류: {sunday_str} 행정광고.pptx 파일을 찾을 수 없습니다.{RESET}")
        print("먼저 행정광고 PPT를 생성하세요.")
        return False

    print(f"{GREEN}PPT 파일 발견: {os.path.basename(ppt_path)}{RESET}")

    # PowerPoint로 PPT 파일 열기
    powerpoint, presentation = open_ppt_in_powerpoint(ppt_path)

    if not powerpoint or not presentation:
        print(f"{RED}PPT 파일을 열 수 없습니다.{RESET}")
        return False

    # PPT 파일 메타데이터 가져오기 (python-pptx 사용)
    try:
        prs = Presentation(ppt_path)
        total_slides = len(prs.slides)
        print(f"총 슬라이드 수: {total_slides}\n")
    except Exception as e:
        print(f"{RED}PPT 파일 정보를 읽을 수 없습니다: {str(e)}{RESET}")
        close_ppt_in_powerpoint(powerpoint, presentation, save=False)
        return False

    # 반복 처리
    while True:
        print("\n" + "=" * 60)
        print(f"{CYAN}텍스트 박스 추가{RESET}")
        print("=" * 60)

        # 슬라이드 번호 입력
        while True:
            try:
                slide_input = input(f"\n슬라이드 번호 (1-{total_slides}): ").strip()
                slide_number = int(slide_input)
                if 1 <= slide_number <= total_slides:
                    break
                else:
                    print(f"{RED}1-{total_slides} 사이의 숫자를 입력해주세요.{RESET}")
            except ValueError:
                print(f"{RED}올바른 숫자를 입력해주세요.{RESET}")

        # 위치 선택
        position_choice = get_position_input()

        # 텍스트 내용 입력
        print(f"\n{CYAN}텍스트 내용을 입력하세요:{RESET}")
        text_content = input("텍스트: ").strip()

        if not text_content:
            print(f"{RED}텍스트가 비어있습니다. 추가를 건너뜁니다.{RESET}")
        else:
            # 텍스트 박스 추가 (win32com 사용)
            success = add_text_box_to_slide_com(presentation, slide_number, position_choice, text_content)

        # 추가 입력 여부 확인
        print(f"\n{CYAN}추가로 텍스트 박스를 추가하시겠습니까?{RESET}")
        print("  1. 예 (계속)")
        print("  2. 아니오 (종료)")

        while True:
            choice = input("선택 (1-2): ").strip()
            if choice == '1':
                # 계속 진행
                break
            elif choice == '2':
                print(f"\n{GREEN}스크립트를 종료합니다.{RESET}")
                print(f"{YELLOW}PowerPoint 창은 열린 상태로 유지됩니다. 직접 확인하세요.{RESET}")
                close_ppt_in_powerpoint(powerpoint, presentation, save=False)
                return True
            else:
                print(f"{RED}1 또는 2를 입력해주세요.{RESET}")


if __name__ == "__main__":
    try:
        success = main()
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print(f"\n\n{YELLOW}프로그램을 종료합니다.{RESET}")
        sys.exit(0)

