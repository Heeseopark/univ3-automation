#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
카카오톡 행정광고 이미지 자동 전송 스크립트
이번 주 일요일 날짜의 행정광고 PNG 폴더 내 이미지들을 카카오톡 나와의 채팅으로 묶어서 전송합니다.
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
import time
from datetime import datetime, timedelta
from pathlib import Path
import pyautogui
import pyperclip
import pytz

# ====================================
# 설정
# ====================================
# 카카오톡 채팅창 이미지 전송 대기 시간 (초)
WAIT_TIME_SHORT = 0.2
WAIT_TIME_MEDIUM = 1
WAIT_TIME_LONG = 1.5

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

def find_png_folder(admin_dir, sunday_str):
    """행정광고 디렉토리에서 이번 주 일요일 PNG 폴더 찾기"""
    png_folder_name = f"{sunday_str} 행정광고 PNG"
    png_folder_path = os.path.join(admin_dir, png_folder_name)

    if os.path.exists(png_folder_path):
        return png_folder_path
    return None

def get_png_files(png_folder):
    """PNG 폴더에서 이미지 파일 목록 가져오기"""
    png_files = []
    for file in os.listdir(png_folder):
        if file.lower().endswith('.png'):
            png_files.append(os.path.join(png_folder, file))

    # 파일명 기준 정렬
    png_files.sort()
    return png_files

def open_kakao_chat_with_myself():
    """카카오톡 나와의 채팅 열기"""
    print("\n" + "="*50)
    print("카카오톡 준비")
    print("="*50)
    print("1. 카카오톡을 실행하세요.")
    print("2. '나와의 채팅' 방으로 들어가세요.")
    print("3. 채팅창을 클릭하여 활성화하세요.")
    print("\n10초 후에 자동으로 이미지 전송이 시작됩니다...")

    for i in range(10, 0, -1):
        print(f"\r{i}초 남음...", end="", flush=True)
        time.sleep(1)

    print("\n\n이미지 전송을 시작합니다.")

def select_images_in_kakao(png_files):
    """여러 이미지를 파일 대화상자에서 전체 선택까지만 진행 (사용자가 제거할 이미지 선택 가능)"""
    print(f"\n총 {len(png_files)}개의 이미지가 있습니다.")
    print("전체 선택 후 멈춥니다. 제거할 이미지를 Ctrl+클릭으로 선택 해제하세요.")

    # 파일 첨부 단축키 (Ctrl+T)
    pyautogui.hotkey('ctrl', 't')
    time.sleep(WAIT_TIME_MEDIUM)

    # PNG 폴더 경로로 이동
    folder_path = os.path.dirname(png_files[0])
    pyperclip.copy(folder_path)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(WAIT_TIME_SHORT)
    pyautogui.press('enter')
    time.sleep(WAIT_TIME_MEDIUM)

    # 파일 목록으로 포커스 이동 (Shift+Tab으로 파일 목록 영역으로 이동)
    pyautogui.hotkey('shift', 'tab')
    time.sleep(WAIT_TIME_SHORT)

    # 모든 파일 선택 (Ctrl+A)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(WAIT_TIME_SHORT)

    print("\n" + "="*50)
    print("전체 이미지가 선택되었습니다!")
    print("="*50)
    print("- 제거할 이미지: Ctrl + 클릭으로 선택 해제")
    print("- 전송하려면: '열기' 버튼 클릭 후 '전송' 버튼 클릭")
    print("="*50)

def main():
    """메인 실행 함수"""
    print("=" * 50)
    print("카카오톡 행정광고 이미지 자동 전송")
    print("=" * 50 + "\n")

    # 디렉토리 설정
    parent_dir = get_base_dir()
    admin_dir = os.path.join(parent_dir, "행정광고")

    if not os.path.exists(admin_dir):
        print(f"오류: {admin_dir} 디렉토리를 찾을 수 없습니다.")
        return False

    # 이번주 일요일 날짜
    this_sunday = get_this_sunday()
    sunday_str = this_sunday.strftime("%m%d")

    print(f"이번주 일요일: {this_sunday.strftime('%Y년 %m월 %d일')} ({sunday_str})")

    # PNG 폴더 찾기
    png_folder = find_png_folder(admin_dir, sunday_str)

    if not png_folder:
        print(f"\n오류: {sunday_str} 행정광고 PNG 폴더를 찾을 수 없습니다.")
        print("먼저 export_and_send_announcement.py를 실행하여 PNG 파일을 생성하세요.")
        return False

    print(f"PNG 폴더 발견: {os.path.basename(png_folder)}")

    # PNG 파일 목록 가져오기
    png_files = get_png_files(png_folder)

    if not png_files:
        print(f"\nPNG 폴더에 이미지 파일이 없습니다.")
        return False

    print(f"발견된 이미지 파일 수: {len(png_files)}개")
    for png_file in png_files:
        print(f"  - {os.path.basename(png_file)}")

    # 30장 초과 체크
    if len(png_files) > 30:
        print("\n" + "="*50)
        print("오류: 이미지가 30장을 초과합니다.")
        print(f"현재 이미지 수: {len(png_files)}장")
        print("카카오톡은 최대 30장까지만 한 번에 전송 가능합니다.")
        print("이미지를 줄여서 다시 시도해주세요.")
        print("="*50)
        return False

    # 자동 전송 안내
    print("\n" + "="*50)
    print("자동 전송 준비")
    print("="*50)
    print("주의사항:")
    print("1. 카카오톡 PC 버전이 설치되어 있어야 합니다.")
    print("2. 카카오톡에 로그인되어 있어야 합니다.")
    print("3. 이미지 전송이 시작되면 마우스나 키보드를 조작하지 마세요.")

    try:
        # 카카오톡 나와의 채팅 열기
        open_kakao_chat_with_myself()

        # 이미지 전체 선택까지만 진행 (사용자가 수동으로 제거 후 전송)
        select_images_in_kakao(png_files)

        print("\n이미지 선택이 완료되었습니다.")
        print("제거할 이미지를 선택 해제한 후 직접 전송해주세요.")
        return True

    except Exception as e:
        print(f"\n오류 발생: {str(e)}")
        print("\n가능한 해결 방법:")
        print("1. 카카오톡 PC 버전이 실행 중인지 확인하세요.")
        print("2. 화면 해상도나 카카오톡 창 위치가 변경되었는지 확인하세요.")
        print("3. pyautogui.displayMousePosition()으로 좌표를 재확인하세요.")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
