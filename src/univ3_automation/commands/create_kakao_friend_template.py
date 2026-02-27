#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
카카오톡 플러스친구 템플릿 자동 생성 스크립트
이번 주 일요일 날짜의 플러스친구 메시지 템플릿을 자동으로 생성합니다.
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

def countdown_timer(seconds, message=""):
    """카운트다운 타이머 표시"""
    for remaining in range(seconds, 0, -1):
        print(f"\r{message} {remaining}초 남음...", end='', flush=True)
        time.sleep(1)
    print("\r" + " " * 50 + "\r", end='', flush=True)

def find_kakao_friend_files(kakao_friend_dir, sunday_date):
    """카톡플친 파일 존재 확인"""
    mmdd = sunday_date.strftime("%m%d")

    # 플러스친구 커버 이미지 찾기
    cover_file = os.path.join(kakao_friend_dir, f"카톡플친 커버 {mmdd}.png")
    if not os.path.exists(cover_file):
        return None, None

    # 예수사람 워터마크 찾기
    watermark_file = os.path.join(kakao_friend_dir, "예수사람 워터마크.png")
    if not os.path.exists(watermark_file):
        return cover_file, None

    return cover_file, watermark_file

def upload_file_with_path(file_path):
    """파일 경로를 직접 입력하여 업로드"""
    time.sleep(0.5)
    pyperclip.copy(file_path)
    pyautogui.hotkey('ctrl', 'v')
    # pyautogui.press('right')
    pyautogui.press('enter')

def create_kakao_friend_template(sunday_date, cover_file, watermark_file):
    """카카오톡 플러스친구 템플릿 생성"""

    # 날짜 정보 생성
    m = sunday_date.strftime("%-m") if os.name != 'nt' else sunday_date.strftime("%#m")
    dd = sunday_date.strftime("%d")

    # 제목 생성
    title = f"[{m}월 {dd}일 예수사람 집회 안내]"

    print(f"\n{CYAN}카카오톡 플러스친구 템플릿 생성을 시작합니다.{RESET}")
    print(f"제목: {MAGENTA}{title}{RESET}")
    print("\n" + "="*60)

    try:
        # 1. Windows 키 누르고 Chrome 실행
        print(f"{GREEN}1. Chrome 브라우저를 실행합니다...{RESET}")
        pyautogui.press('win')
        pyautogui.write('chrome')
        pyautogui.press('enter')
        time.sleep(2)  # Chrome 로딩 대기

        # 2. 카카오톡 플러스친구 URL 입력
        print(f"{GREEN}2. 카카오톡 플러스친구 관리자 페이지로 이동합니다...{RESET}")
        url = "https://business.kakao.com/_lcxnrd/messages/new/widelist"
        pyautogui.hotkey('ctrl', 'l')  # 주소창 포커스
        time.sleep(0.5)
        pyautogui.write(url)
        pyautogui.press('enter')

        # 3. 로그인 및 제목 입력 대기
        print(f"\n{YELLOW}3. 로그인이 필요한 경우 로그인해주세요.{RESET}")
        print(f"{YELLOW}   그리고 제목 입력란에 커서를 두세요.{RESET}")
        countdown_timer(10, "대기 중:")

        # 4. 제목 입력
        print(f"{GREEN}4. 제목을 입력합니다...{RESET}")
        pyperclip.copy(title)
        pyautogui.hotkey('ctrl', 'v')

        # 5. Tab 3번, 엔터
        print(f"{GREEN}5. 다음 섹션으로 이동합니다...{RESET}")
        for _ in range(3):
            pyautogui.press('tab')
        pyautogui.press('enter')

        # 6. 플러스친구 커버 이미지 업로드
        print(f"{GREEN}6. 플러스친구 커버 이미지를 업로드합니다...{RESET}")
        upload_file_with_path(cover_file)

        # 7. Tab 2번, 텍스트 입력
        print(f"{GREEN}7. 첫 번째 링크 텍스트를 입력합니다...{RESET}")
        time.sleep(2)
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyperclip.copy("☞ 클릭하면 주보로!")
        pyautogui.hotkey('ctrl', 'v')

        # 8. Tab 6번, 엔터, 워터마크 첨부
        print(f"{GREEN}8. 첫 번째 워터마크를 첨부합니다...{RESET}")
        for _ in range(6):
            pyautogui.press('tab')
        pyautogui.press('enter')
        upload_file_with_path(watermark_file)

        # 9. Tab 2번, 텍스트 입력
        print(f"{GREEN}9. 두 번째 링크 텍스트를 입력합니다...{RESET}")
        time.sleep(2)
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyperclip.copy("☞ ")
        pyautogui.hotkey('ctrl', 'v')

        # 10. Tab 6번, 엔터, 워터마크 첨부
        print(f"{GREEN}10. 두 번째 워터마크를 첨부합니다...{RESET}")
        for _ in range(6):
            pyautogui.press('tab')
        pyautogui.press('enter')
        upload_file_with_path(watermark_file)

        # 11. Tab 2번, 텍스트 입력
        print(f"{GREEN}11. 세 번째 링크 텍스트를 입력합니다...{RESET}")
        time.sleep(2)
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyperclip.copy("☞ ")
        pyautogui.hotkey('ctrl', 'v')

        # 12. Tab 4번, 엔터
        print(f"{GREEN}12. 다음 섹션으로 이동합니다...{RESET}")
        for _ in range(4):
            pyautogui.press('tab')
        pyautogui.press('enter')

        # 13. Tab 3번, 엔터, 워터마크 첨부
        print(f"{GREEN}13. 마지막 워터마크를 첨부합니다...{RESET}")
        for _ in range(3):
            pyautogui.press('tab')
        pyautogui.press('enter')
        pyautogui.press('enter')
        upload_file_with_path(watermark_file)

        # 14. Tab 2번, 텍스트 입력
        print(f"{GREEN}14. 마지막 링크 텍스트를 입력합니다...{RESET}")
        time.sleep(2)
        for _ in range(2):
            pyautogui.press('tab')
        pyperclip.copy("☞ ")
        pyautogui.hotkey('ctrl', 'v')

        print(f"\n{GREEN}✓ 카카오톡 플러스친구 템플릿 생성이 완료되었습니다!{RESET}")
        print("\n" + "="*60)
        print(f"{CYAN}이제 다음 작업을 수행해주세요:{RESET}")
        print(f"  1. 링크 URL을 입력하세요")
        print(f"  2. 템플릿을 저장하세요")
        print("\n" + "="*60)

        return True

    except Exception as e:
        print(f"\n{RED}오류 발생: {str(e)}{RESET}")
        return False

def main():
    """메인 실행 함수"""
    clear_screen()
    print("=" * 60)
    print(f"{CYAN}       카카오톡 플러스친구 템플릿 생성{RESET}")
    print("=" * 60)

    # pyautogui 안전장치 설정
    pyautogui.FAILSAFE = True  # 화면 모서리로 마우스 이동 시 중단
    pyautogui.PAUSE = 0.5  # 각 명령 사이 0.5초 대기

    # 이번주 일요일 날짜
    this_sunday = get_this_sunday()
    sunday_str = this_sunday.strftime("%m%d")

    print(f"\n이번주 일요일: {MAGENTA}{this_sunday.strftime('%Y년 %m월 %d일')} ({sunday_str}){RESET}")

    # 카톡플친 파일 확인
    parent_dir = get_base_dir()
    kakao_friend_dir = os.path.join(parent_dir, "카톡플친")

    cover_file, watermark_file = find_kakao_friend_files(kakao_friend_dir, this_sunday)

    if not cover_file:
        print(f"\n{RED}오류: {sunday_str} 날짜의 플러스친구 커버 이미지를 찾을 수 없습니다.{RESET}")
        print(f"파일명: 카톡플친 커버 {sunday_str}.png")
        print(f"위치: {kakao_friend_dir}")
        return False

    if not watermark_file:
        print(f"\n{RED}오류: 예수사람 워터마크 파일을 찾을 수 없습니다.{RESET}")
        print(f"위치: {kakao_friend_dir}")
        return False

    print(f"\n커버 이미지: {GREEN}{os.path.basename(cover_file)}{RESET}")
    print(f"워터마크: {GREEN}{os.path.basename(watermark_file)}{RESET}")

    # 업로드 시작 확인
    print(f"\n{CYAN}카카오톡 플러스친구 템플릿 생성을 시작하시겠습니까?{RESET}")
    print("Enter를 누르면 시작, ESC를 누르면 취소")
    print(f"\n{YELLOW}주의: 자동화 중 마우스나 키보드를 건드리지 마세요!{RESET}")
    print(f"{YELLOW}중단하려면 마우스를 화면 모서리로 이동하세요.{RESET}")

    if os.name == 'nt':
        import msvcrt
        key = msvcrt.getch()
        if key == b'\x1b':  # ESC
            print(f"\n{YELLOW}작업이 취소되었습니다.{RESET}")
            return False
    else:
        response = input()
        if response.lower() == 'n':
            print(f"\n{YELLOW}작업이 취소되었습니다.{RESET}")
            return False

    # 카카오톡 플러스친구 템플릿 생성 실행
    success = create_kakao_friend_template(this_sunday, cover_file, watermark_file)

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

