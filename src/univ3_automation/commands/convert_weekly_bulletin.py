#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
주보 PPT PDF/PNG 변환 스크립트
이번 주 일요일 주보를 PDF와 PNG로 변환합니다.
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
import shutil
from datetime import datetime, timedelta
import pytz
import win32com.client
import pythoncom
from pathlib import Path
import glob

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
except ImportError:
    MAGENTA = CYAN = WHITE = RESET = GREEN = YELLOW = ''

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

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

def find_bulletin_ppt(bulletin_dir, sunday_date):
    """주보 PPT 파일 찾기"""
    mmdd = sunday_date.strftime("%m%d")

    # 파일 패턴으로 검색 (호수는 무시하고 날짜만으로 검색)
    pattern = f"*{mmdd} 주보.pptx"
    all_files = glob.glob(os.path.join(bulletin_dir, pattern))

    # 임시 파일 (~$로 시작하는 파일) 제외
    files = [f for f in all_files if not os.path.basename(f).startswith("~$")]

    if files:
        # 가장 최신 파일 선택 (여러 개가 있을 경우)
        return max(files, key=os.path.getmtime)

    return None

def export_ppt_to_pdf_and_png(ppt_path, output_dir):
    """PPT를 PDF와 PNG로 변환"""
    print(f"\n{CYAN}PPT 파일 변환 시작: {os.path.basename(ppt_path)}{RESET}")

    # 파일이 임시 파일인지 재확인
    if os.path.basename(ppt_path).startswith("~$"):
        print(f"{YELLOW}경고: 임시 파일은 변환할 수 없습니다. 정상 파일을 사용하세요.{RESET}")
        return None, None

    pythoncom.CoInitialize()

    try:
        # PowerPoint 애플리케이션 시작
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1  # PowerPoint는 반드시 Visible이어야 함
        powerpoint.WindowState = 2  # ppWindowMinimized - 최소화 상태로 실행

        # PPT 파일 열기
        abs_ppt_path = os.path.abspath(ppt_path)
        presentation = powerpoint.Presentations.Open(abs_ppt_path, ReadOnly=True)

        # PDF 경로 설정
        base_name = os.path.basename(ppt_path).replace('.pptx', '')
        pdf_path = os.path.join(output_dir, f"{base_name} PDF.pdf")

        # 기존 PDF 파일이 있으면 삭제
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
            print(f"기존 PDF 파일 삭제: {os.path.basename(pdf_path)}")

        # PDF로 저장
        print(f"{GREEN}PDF 생성 중...{RESET}")
        presentation.SaveAs(pdf_path, 32)  # 32 = ppSaveAsPDF
        print(f"{GREEN}✓ PDF 생성 완료: {os.path.basename(pdf_path)}{RESET}")

        # PNG 디렉토리 생성
        png_dir = os.path.join(output_dir, f"{base_name} PNG")
        if os.path.exists(png_dir):
            try:
                shutil.rmtree(png_dir)
                print(f"기존 PNG 디렉토리 삭제: {os.path.basename(png_dir)}")
            except Exception as e:
                print(f"{YELLOW}기존 디렉토리 삭제 실패, 덮어쓰기 시도: {e}{RESET}")

        # 디렉토리가 없을 때만 생성
        if not os.path.exists(png_dir):
            os.makedirs(png_dir)

        # 슬라이드 크기 가져오기
        slide_width = presentation.PageSetup.SlideWidth
        slide_height = presentation.PageSetup.SlideHeight

        # 모든 슬라이드를 PNG로 저장
        total_slides = presentation.Slides.Count
        print(f"\n{GREEN}PNG 추출 중 (전체 {total_slides}개 슬라이드)...{RESET}")

        for i in range(1, total_slides + 1):
            # Item() 메서드를 사용하여 슬라이드에 접근 (더 안전한 방법)
            slide = presentation.Slides.Item(i)
            png_file = os.path.join(png_dir, f"슬라이드{i:02d}.png")

            # 슬라이드를 명시적으로 선택한 후 Export
            slide.Select()
            slide.Export(png_file, "PNG", int(slide_width * 2), int(slide_height * 2))  # 고해상도로 저장
            print(f"  {CYAN}✓ 슬라이드 {i}/{total_slides} 추출 완료{RESET}")

        print(f"{GREEN}✓ PNG 추출 완료: {os.path.basename(png_dir)}{RESET}")

        # PowerPoint 종료
        presentation.Close()
        powerpoint.Quit()

        return pdf_path, png_dir

    except Exception as e:
        print(f"{YELLOW}변환 중 오류 발생: {str(e)}{RESET}")
        import traceback
        traceback.print_exc()
        return None, None

    finally:
        pythoncom.CoUninitialize()

def main():
    """메인 실행 함수"""
    clear_screen()
    print("=" * 60)
    print(f"{CYAN}       주보 PPT → PDF/PNG 변환{RESET}")
    print("=" * 60)

    # 디렉토리 설정
    parent_dir = get_base_dir()
    bulletin_dir = os.path.join(parent_dir, "주보")

    if not os.path.exists(bulletin_dir):
        print(f"{YELLOW}오류: {bulletin_dir} 디렉토리를 찾을 수 없습니다.{RESET}")
        return False

    # 이번주 일요일 날짜
    this_sunday = get_this_sunday()
    sunday_str = this_sunday.strftime("%m%d")

    print(f"\n이번주 일요일: {MAGENTA}{this_sunday.strftime('%Y년 %m월 %d일')} ({sunday_str}){RESET}")

    # 주보 PPT 파일 찾기
    ppt_path = find_bulletin_ppt(bulletin_dir, this_sunday)

    if not ppt_path:
        print(f"\n{YELLOW}오류: {sunday_str} 날짜의 주보 PPT 파일을 찾을 수 없습니다.{RESET}")
        print(f"찾는 패턴: *{sunday_str} 주보.pptx")
        print(f"주보 디렉토리: {bulletin_dir}")

        # 주보 디렉토리의 파일 목록 보여주기
        print(f"\n{CYAN}현재 주보 디렉토리의 파일들:{RESET}")
        pptx_files = glob.glob(os.path.join(bulletin_dir, "*.pptx"))
        if pptx_files:
            for f in sorted(pptx_files)[-5:]:  # 최근 5개만 표시
                print(f"  - {os.path.basename(f)}")
        else:
            print("  (PPT 파일이 없습니다)")

        return False

    print(f"{GREEN}PPT 파일 발견: {os.path.basename(ppt_path)}{RESET}")

    # 변환 시작 확인
    print(f"\n{CYAN}이 파일을 PDF와 PNG로 변환하시겠습니까?{RESET}")
    print("Enter를 누르면 시작, ESC를 누르면 취소")

    if os.name == 'nt':
        import msvcrt
        key = msvcrt.getch()
        if key == b'\x1b':  # ESC
            print(f"\n{YELLOW}변환이 취소되었습니다.{RESET}")
            return False
    else:
        response = input()
        if response.lower() == 'n':
            print(f"\n{YELLOW}변환이 취소되었습니다.{RESET}")
            return False

    # PDF와 PNG로 변환
    pdf_path, png_dir = export_ppt_to_pdf_and_png(ppt_path, bulletin_dir)

    if not pdf_path:
        print(f"\n{YELLOW}변환에 실패했습니다.{RESET}")
        return False

    print("\n" + "=" * 60)
    print(f"{GREEN}       변환 완료!{RESET}")
    print("=" * 60)
    print(f"\n{CYAN}생성된 파일:{RESET}")
    print(f"  PDF: {os.path.basename(pdf_path)}")
    print(f"  PNG: {os.path.basename(png_dir)}/ ({len(os.listdir(png_dir))}개 파일)")
    print(f"\n위치: {bulletin_dir}")

    return True

if __name__ == "__main__":
    success = main()

    print(f"\n{CYAN}아무 키나 누르면 종료합니다...{RESET}")
    if os.name == 'nt':
        import msvcrt
        msvcrt.getch()
    else:
        input()

    sys.exit(0 if success else 1)
