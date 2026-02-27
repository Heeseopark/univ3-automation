#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
카카오톡 행정광고 PPT 자동 전송 스크립트
이번 주 일요일 날짜의 행정광고 PPT 파일을 카카오톡 나와의 채팅으로 전송합니다.
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
# 카카오톡 채팅창 파일 전송 대기 시간 (초)
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

def find_admin_ppt(admin_dir, sunday_str):
    """행정광고 디렉토리에서 이번 주 일요일 PPT 파일 찾기"""
    ppt_filename = f"{sunday_str} 행정광고.pptx"
    ppt_path = os.path.join(admin_dir, ppt_filename)

    if os.path.exists(ppt_path):
        return ppt_path
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

    print("\n\nPPT 파일 전송을 시작합니다.")

def send_ppt_to_kakao(ppt_path):
    """PPT 파일을 카카오톡으로 전송"""
    print(f"\nPPT 파일 전송 중: {os.path.basename(ppt_path)}")

    # 파일 첨부 단축키 (Ctrl+T)
    pyautogui.hotkey('ctrl', 't')
    time.sleep(WAIT_TIME_MEDIUM)

    # 파일 경로 입력
    pyperclip.copy(ppt_path)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(WAIT_TIME_SHORT)
    pyautogui.press('enter')
    time.sleep(WAIT_TIME_MEDIUM)

    # 전송 버튼 클릭 (Enter)
    pyautogui.press('enter')
    time.sleep(WAIT_TIME_LONG)

    print("PPT 파일 전송 완료!")

def main():
    """메인 실행 함수"""
    print("=" * 50)
    print("카카오톡 행정광고 PPT 자동 전송")
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

    # PPT 파일 찾기
    ppt_path = find_admin_ppt(admin_dir, sunday_str)

    if not ppt_path:
        print(f"\n오류: {sunday_str} 행정광고.pptx 파일을 찾을 수 없습니다.")
        print("먼저 '행정광고 PPT 생성' 기능을 실행하세요.")
        return False

    print(f"PPT 파일 발견: {os.path.basename(ppt_path)}")
    print(f"파일 크기: {os.path.getsize(ppt_path) / 1024 / 1024:.2f} MB")

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

        # PPT 파일 전송
        send_ppt_to_kakao(ppt_path)

        print("\n모든 작업이 성공적으로 완료되었습니다!")
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

