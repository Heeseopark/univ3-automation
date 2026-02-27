#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
주보 카카오톡 채널 업로드 스크립트
이번 주 일요일 주보를 카카오톡 채널에 자동으로 업로드합니다.
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
import pytz
import pyautogui
import pyperclip
import subprocess
from pathlib import Path

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

def check_bulletin_files(bulletin_dir, sunday_date):
    """주보 파일 존재 확인"""
    mmdd = sunday_date.strftime("%m%d")

    # PNG 디렉토리 찾기
    import glob
    pattern = f"*{mmdd} 주보 PNG"
    png_dirs = glob.glob(os.path.join(bulletin_dir, pattern))

    # 임시 파일 제외
    png_dirs = [d for d in png_dirs if os.path.isdir(d) and not os.path.basename(d).startswith("~")]

    if not png_dirs:
        return None

    # 가장 최근 디렉토리 선택
    png_dir = max(png_dirs, key=os.path.getmtime)

    # PNG 파일들 확인
    png_files = sorted(glob.glob(os.path.join(png_dir, "*.png")))

    return png_dir, png_files

def countdown_timer(seconds, message=""):
    """카운트다운 타이머 표시"""
    for remaining in range(seconds, 0, -1):
        print(f"\r{message} {remaining}초 남음...", end='', flush=True)
        time.sleep(1)
    print("\r" + " " * 50 + "\r", end='', flush=True)

def upload_to_kakao_channel(sunday_date, png_dir):
    """카카오톡 채널에 주보 업로드"""

    # 날짜 정보 생성
    yyyy = sunday_date.strftime("%Y")
    m = sunday_date.strftime("%-m") if os.name != 'nt' else sunday_date.strftime("%#m")
    dd = sunday_date.strftime("%d")

    # 제목 생성
    title = f"[대학3부 예수사람] 주후 {yyyy}년 {m}월 {dd}일 주보"
    description = f"주후 {yyyy}년 {m}월 {dd}일 주보입니다."

    print(f"\n{CYAN}카카오톡 채널 업로드를 시작합니다.{RESET}")
    print(f"제목: {MAGENTA}{title}{RESET}")
    print("\n" + "="*60)

    try:
        # 1. Windows 키 누르고 Chrome 실행
        print(f"{GREEN}1. Chrome 브라우저를 실행합니다...{RESET}")
        pyautogui.press('win')
        time.sleep(0.5)
        pyautogui.write('chrome')
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(2)  # Chrome 로딩 대기

        # 2. 카카오톡 채널 URL 입력
        print(f"{GREEN}2. 카카오톡 채널 관리자 페이지로 이동합니다...{RESET}")
        url = "https://business.kakao.com/_lcxnrd/posts?t_src=business_partnercenter&t_ch=lnb"
        pyautogui.hotkey('ctrl', 'l')  # 주소창 포커스
        pyperclip.copy(url)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')

        # 3. 로그인 대기
        print(f"\n{YELLOW}3. 로그인이 필요한 경우 로그인해주세요.{RESET}")
        countdown_timer(7, "대기 중:")

        # 4. 포스트 작성 단축키 입력 대기
        print(f"{GREEN}4. 포스트 작성 준비 중...{RESET}")
        print(f"\n{YELLOW}특별한 포커스 없이 소식 페이지를 열어주세요.{RESET}")
        countdown_timer(5, "대기 중:")

        # 5. 포스트 작성 단축키 입력
        print(f"{GREEN}5. 포스트 작성 창을 엽니다...{RESET}")
        pyautogui.press('tab')

        # 6. 제목 입력
        print(f"{GREEN}6. 제목을 입력합니다...{RESET}")
        pyperclip.copy(title)
        time.sleep(0.2)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)

        print(f"{GREEN}7. 설명을 입력합니다...{RESET}")
        pyautogui.press('tab')
        pyperclip.copy(description)
        time.sleep(0.2)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)

        # 7. 내용 작성 영역으로 이동 및 이미지 준비
        print(f"{GREEN}8. 이미지 업로드를 준비합니다...{RESET}")
        pyautogui.press('tab')
        pyautogui.press('tab')

        # 8. 엔터 입력하여 파일 선택 창 열기
        print(f"{GREEN}9. 이미지 업로드 창을 엽니다...{RESET}")
        pyautogui.press('enter')
        time.sleep(1)

        # 9. PNG 폴더 경로 입력
        print(f"{GREEN}10. PNG 폴더 경로를 입력합니다...{RESET}")
        pyperclip.copy(png_dir)
        time.sleep(0.2)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.2)
        pyautogui.press('enter')
        time.sleep(1)

        # 10. 파일 목록으로 포커스 이동 (Shift+Tab 키로 이동)
        print(f"{GREEN}11. 파일 목록으로 이동합니다...{RESET}")
        pyautogui.hotkey('shift', 'tab')
        time.sleep(0.2)

        # 11. 모든 파일 선택 (Ctrl+A)
        print(f"{GREEN}12. 모든 PNG 파일을 선택합니다...{RESET}")
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.2)

        # 12. 선택한 파일들 열기
        print(f"{GREEN}13. 파일을 업로드합니다...{RESET}")
        pyautogui.press('enter')
        time.sleep(1.5)

        # 13. 전송 버튼 클릭 (Enter)
        print(f"{GREEN}14. 업로드를 확정합니다...{RESET}")
        pyautogui.press('enter')
        time.sleep(1)

        print(f"\n{GREEN}✓ 카카오톡 채널 업로드가 완료되었습니다!{RESET}")
        print("\n" + "="*60)
        print(f"{CYAN}이제 다음 작업을 수행해주세요:{RESET}")
        print(f"  1. 업로드된 이미지를 확인하세요")
        print(f"  2. 포스트를 발행하세요")
        print("\n" + "="*60)

        return True

    except Exception as e:
        print(f"\n{RED}오류 발생: {str(e)}{RESET}")
        return False

def main():
    """메인 실행 함수"""
    clear_screen()
    print("=" * 60)
    print(f"{CYAN}       주보 카카오톡 채널 업로드{RESET}")
    print("=" * 60)

    # pyautogui 안전장치 설정
    pyautogui.FAILSAFE = True  # 화면 모서리로 마우스 이동 시 중단
    pyautogui.PAUSE = 0.5  # 각 명령 사이 0.5초 대기

    # 이번주 일요일 날짜
    this_sunday = get_this_sunday()
    sunday_str = this_sunday.strftime("%m%d")

    print(f"\n이번주 일요일: {MAGENTA}{this_sunday.strftime('%Y년 %m월 %d일')} ({sunday_str}){RESET}")

    # 주보 파일 확인
    parent_dir = get_base_dir()
    bulletin_dir = os.path.join(parent_dir, "주보")

    result = check_bulletin_files(bulletin_dir, this_sunday)

    if not result:
        print(f"\n{YELLOW}경고: {sunday_str} 날짜의 주보 PNG 파일을 찾을 수 없습니다.{RESET}")
        print(f"먼저 주보 PPT를 PDF/PNG로 변환해주세요.")
        return False

    png_dir, png_files = result
    print(f"\nPNG 디렉토리: {GREEN}{os.path.basename(png_dir)}{RESET}")
    print(f"PNG 파일 수: {GREEN}{len(png_files)}개{RESET}")

    # 업로드 시작 확인
    print(f"\n{CYAN}카카오톡 채널에 업로드를 시작하시겠습니까?{RESET}")
    print("Enter를 누르면 시작, ESC를 누르면 취소")
    print(f"\n{YELLOW}주의: 자동화 중 마우스나 키보드를 건드리지 마세요!{RESET}")
    print(f"{YELLOW}중단하려면 마우스를 화면 모서리로 이동하세요.{RESET}")

    if os.name == 'nt':
        import msvcrt
        key = msvcrt.getch()
        if key == b'\x1b':  # ESC
            print(f"\n{YELLOW}업로드가 취소되었습니다.{RESET}")
            return False
    else:
        response = input()
        if response.lower() == 'n':
            print(f"\n{YELLOW}업로드가 취소되었습니다.{RESET}")
            return False

    # 카카오톡 채널 업로드 실행
    success = upload_to_kakao_channel(this_sunday, png_dir)

    if success:
        print(f"\n{GREEN}PNG 파일 위치:{RESET}")
        print(f"{CYAN}{png_dir}{RESET}")

    return success

if __name__ == "__main__":
    success = main()

    print(f"\n{CYAN}아무 키나 누르면 종료합니다...{RESET}")
    if os.name == 'nt':
        import msvcrt
        msvcrt.getch()
    else:
        input()

    sys.exit(0 if success else 1)
