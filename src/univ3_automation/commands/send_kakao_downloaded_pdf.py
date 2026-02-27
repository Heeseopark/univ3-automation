#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
다운로드 폴더의 최근 PDF 파일을 카카오톡 나와의 채팅으로 전송
마지막 5개의 PDF 파일을 보여주고 선택할 수 있게 함
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
from pathlib import Path
from datetime import datetime
import pyautogui
import pyperclip

# 카카오톡 파일 전송 대기 시간 (초)
WAIT_TIME_SHORT = 0.2
WAIT_TIME_MEDIUM = 1
WAIT_TIME_LONG = 1.5

# 키보드 입력 처리
try:
    import msvcrt

    def get_key():
        """Windows에서 키 입력 받기"""
        key = msvcrt.getch()
        if key in [b'\x00', b'\xe0']:  # 특수 키
            key = msvcrt.getch()
            if key == b'H':  # 위 화살표
                return 'UP'
            elif key == b'P':  # 아래 화살표
                return 'DOWN'
        elif key == b'\r':  # Enter
            return 'ENTER'
        elif key == b'q' or key == b'Q':  # Q
            return 'Q'
        return None

except ImportError:
    import tty, termios

    def get_key():
        """Unix/Linux/Mac에서 키 입력 받기"""
        fd = sys.stdin.fileno()
        old_settings = termios.tcgetattr(fd)
        try:
            tty.setraw(sys.stdin.fileno())
            key = sys.stdin.read(1)
            if key == '\x1b':  # ESC sequence
                key += sys.stdin.read(2)
                if key == '\x1b[A':  # 위 화살표
                    return 'UP'
                elif key == '\x1b[B':  # 아래 화살표
                    return 'DOWN'
            elif key == '\r' or key == '\n':  # Enter
                return 'ENTER'
            elif key == 'q' or key == 'Q':  # Q
                return 'Q'
        finally:
            termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)
        return None

def clear_screen():
    """화면 초기화"""
    os.system('cls' if os.name == 'nt' else 'clear')

def get_recent_pdfs(download_dir, count=5):
    """다운로드 폴더에서 최근 PDF 파일 찾기"""
    if not os.path.exists(download_dir):
        return []

    # PDF 파일 목록 가져오기
    pdf_files = list(Path(download_dir).glob("*.pdf"))

    # 수정 시간 기준으로 정렬 (최신순)
    pdf_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)

    return pdf_files[:count]

def display_pdf_list(pdf_files):
    """PDF 파일 목록 표시"""
    print("\n최근 다운로드한 PDF 파일 (최신순):")
    print("=" * 80)

    for i, pdf_file in enumerate(pdf_files, 1):
        file_size = pdf_file.stat().st_size / 1024  # KB
        file_time = datetime.fromtimestamp(pdf_file.stat().st_mtime)
        print(f"{i}. {pdf_file.name}")
        print(f"   크기: {file_size:.1f} KB | 수정 시간: {file_time.strftime('%Y-%m-%d %H:%M:%S')}")
        print()

def select_pdf(pdf_files):
    """전송할 PDF 파일 선택 - 화살표 키로 이동"""
    current_index = 0

    while True:
        clear_screen()
        print("=" * 80)
        print("전송할 파일을 선택하세요 (화살표: 이동, Enter: 선택, Q: 취소)")
        print("=" * 80)
        print()

        for i, pdf_file in enumerate(pdf_files):
            file_size = pdf_file.stat().st_size / 1024  # KB
            file_time = datetime.fromtimestamp(pdf_file.stat().st_mtime)

            prefix = "▶ " if i == current_index else "  "

            print(f"{prefix}{i+1}. {pdf_file.name}")
            print(f"     크기: {file_size:.1f} KB | 수정 시간: {file_time.strftime('%Y-%m-%d %H:%M:%S')}")
            print()

        key = get_key()

        if key == 'UP':
            current_index = (current_index - 1) % len(pdf_files)
        elif key == 'DOWN':
            current_index = (current_index + 1) % len(pdf_files)
        elif key == 'ENTER':
            return pdf_files[current_index]
        elif key == 'Q':
            return None

def open_kakao_chat_with_myself():
    """카카오톡 나와의 채팅 열기"""
    print("\n" + "="*50)
    print("카카오톡 준비")
    print("="*50)
    print("1. 카카오톡을 실행하세요.")
    print("2. '나와의 채팅' 방으로 들어가세요.")
    print("3. 채팅창을 클릭하여 활성화하세요.")
    print("\n10초 후에 자동으로 파일 전송이 시작됩니다...")

    for i in range(10, 0, -1):
        print(f"\r{i}초 남음...", end="", flush=True)
        time.sleep(1)

    print("\n\nPDF 파일 전송을 시작합니다.")

def send_pdf_to_kakao(pdf_path):
    """PDF 파일을 카카오톡으로 전송"""
    print(f"\nPDF 파일 전송 중: {pdf_path.name}")

    # 파일 첨부 단축키 (Ctrl+T)
    pyautogui.hotkey('ctrl', 't')
    time.sleep(WAIT_TIME_MEDIUM)

    # 파일 경로 입력
    pyperclip.copy(str(pdf_path))
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(WAIT_TIME_SHORT)
    pyautogui.press('enter')
    time.sleep(WAIT_TIME_MEDIUM)

    # 전송 버튼 클릭 (Enter)
    pyautogui.press('enter')
    time.sleep(WAIT_TIME_LONG)

    print("PDF 파일 전송 완료!")

def select_source_folder():
    """소스 폴더 선택 - 화살표 키로 이동"""
    folders = [
        ("다운로드 폴더", get_downloads_dir()),
        ("카카오톡 받은 파일", get_kakao_downloads_dir())
    ]
    current_index = 0

    while True:
        clear_screen()
        print("=" * 80)
        print("소스 폴더를 선택하세요 (화살표: 이동, Enter: 선택)")
        print("=" * 80)
        print()

        for i, (name, path) in enumerate(folders):
            prefix = "▶ " if i == current_index else "  "
            print(f"{prefix}{i+1}. {name}")
            print(f"     {path}")
            print()

        key = get_key()

        if key == 'UP':
            current_index = (current_index - 1) % len(folders)
        elif key == 'DOWN':
            current_index = (current_index + 1) % len(folders)
        elif key == 'ENTER':
            return folders[current_index][1]

def main():
    """메인 함수"""
    print("=" * 80)
    print("카카오톡 PDF 파일 자동 전송")
    print("=" * 80)

    # 소스 폴더 선택
    download_dir = select_source_folder()

    print(f"\n소스 폴더: {download_dir}")

    # 최근 PDF 파일 찾기
    pdf_files = get_recent_pdfs(download_dir, count=5)

    if not pdf_files:
        print("\n다운로드 폴더에 PDF 파일이 없습니다.")
        return False

    # PDF 목록 표시
    display_pdf_list(pdf_files)

    # 파일 선택
    selected_pdf = select_pdf(pdf_files)

    if not selected_pdf:
        print("\n작업을 취소했습니다.")
        return False

    print(f"\n선택한 파일: {selected_pdf.name}")

    # 자동 전송 안내
    print("\n" + "="*50)
    print("자동 전송 준비")
    print("="*50)
    print("주의사항:")
    print("1. 카카오톡 PC 버전이 설치되어 있어야 합니다.")
    print("2. 카카오톡에 로그인되어 있어야 합니다.")
    print("3. 파일 전송이 시작되면 마우스나 키보드를 조작하지 마세요.")

    try:
        # 카카오톡 나와의 채팅 열기
        open_kakao_chat_with_myself()

        # PDF 파일 전송
        send_pdf_to_kakao(selected_pdf)

        print("\n모든 작업이 성공적으로 완료되었습니다!")
        return True

    except Exception as e:
        print(f"\n오류 발생: {str(e)}")
        print("\n가능한 해결 방법:")
        print("1. 카카오톡 PC 버전이 실행 중인지 확인하세요.")
        print("2. 화면 해상도나 카카오톡 창 위치가 변경되었는지 확인하세요.")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)

