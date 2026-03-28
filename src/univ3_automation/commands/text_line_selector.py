#!/usr/bin/env python3
# -*- coding: utf-8 -*-

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
import subprocess
from datetime import datetime, timedelta

# 색상 코드 정의
try:
    import colorama
    colorama.init()
    MAGENTA = '\033[95m'
    CYAN = '\033[96m'
    WHITE = '\033[97m'
    RESET = '\033[0m'
except ImportError:
    MAGENTA = CYAN = WHITE = RESET = ''

# 현재 달과 다음 달 계산
now = datetime.now()
current_month = now.month
next_month = (now + timedelta(days=32)).month

# 텍스트 데이터를 배열로 관리 (두 줄씩 묶어서 4개 항목)
TEXT_ITEMS = [
    "❤‍🔥느헤미야 교재 입금 계좌❤‍🔥\n3333096478189 카카오뱅크 박희서",
    f"❤‍🔥{current_month}월 큐티 입금 계좌❤‍🔥\n3333312512962 카카오뱅크 박희서",
    f"❤‍🔥{next_month}월 큐티 입금 계좌❤‍🔥\n3333312512962 카카오뱅크 박희서",
    "❤‍🔥중보기도 요청 구글폼❤‍🔥\nhttps://forms.gle/fW6b71KkWYWbfCvY7",
    "❤‍🔥국외선교 후원 계좌❤‍🔥\n3333350028714 카카오뱅크 안혜경",
    "❤‍🔥선사마 후원 계좌❤‍🔥\n3333265396395 카카오뱅크 오현은"
]

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def copy_to_clipboard(text):
    try:
        # PowerShell을 사용하여 UTF-8 텍스트를 올바르게 클립보드에 복사
        subprocess.run(
            ['powershell', '-Command', f'Set-Clipboard -Value @"\n{text}\n"@'],
            check=True,
            shell=False
        )
        return True
    except:
        return False

def display_menu(items, selected, current_index):
    clear_screen()
    print("=" * 60)
    print(f"{CYAN}텍스트 항목 선택기{RESET}")
    print("=" * 60)
    print("↑/↓: 이동, 스페이스: 선택/해제, Enter: 완료, ESC: 취소")
    print("=" * 60)
    print()

    for i, item in enumerate(items):
        lines = item.split('\n')

        # 체크박스 표시
        selected_mark = f"{MAGENTA}[✓]{RESET}" if i in selected else "[ ]"

        # 커서 표시
        cursor = "▶ " if i == current_index else "  "

        # 선택 여부와 현재 위치에 따라 색상 적용
        if i in selected:
            # 선택된 항목은 마젠타로 표시
            print(f"{cursor}{selected_mark} {MAGENTA}{lines[0]}{RESET}")
            print(f"     {MAGENTA}{lines[1]}{RESET}")
        elif i == current_index:
            # 현재 커서 위치는 시안(cyan)으로 표시
            print(f"{cursor}{selected_mark} {CYAN}{lines[0]}{RESET}")
            print(f"     {CYAN}{lines[1]}{RESET}")
        else:
            # 일반 항목
            print(f"{cursor}{selected_mark} {lines[0]}")
            print(f"     {lines[1]}")
        print()

    print("=" * 60)
    print(f"선택된 항목: {len(selected)}개")

def main():
    items = TEXT_ITEMS

    if not items:
        print("텍스트 데이터가 없습니다.")
        return

    selected = set()
    current_index = 0

    try:
        if os.name == 'nt':
            import msvcrt

            while True:
                display_menu(items, selected, current_index)

                key = msvcrt.getch()

                if key == b'\xe0' or key == b'\x00':
                    key = msvcrt.getch()
                    if key == b'H':  # 위쪽 화살표
                        current_index = max(0, current_index - 1)
                    elif key == b'P':  # 아래쪽 화살표
                        current_index = min(len(items) - 1, current_index + 1)

                elif key == b' ':  # 스페이스바
                    if current_index in selected:
                        selected.remove(current_index)
                    else:
                        selected.add(current_index)

                elif key == b'\r':  # Enter
                    break

                elif key == b'\x1b':  # ESC
                    print("\n작업이 취소되었습니다.")
                    return

        else:
            import termios, tty

            old_settings = termios.tcgetattr(sys.stdin)
            try:
                tty.setraw(sys.stdin.fileno())

                while True:
                    display_menu(items, selected, current_index)

                    key = sys.stdin.read(1)

                    if key == '\x1b':
                        next_key = sys.stdin.read(2)
                        if next_key == '[A':  # 위쪽 화살표
                            current_index = max(0, current_index - 1)
                        elif next_key == '[B':  # 아래쪽 화살표
                            current_index = min(len(items) - 1, current_index + 1)
                        elif next_key == '':
                            print("\n작업이 취소되었습니다.")
                            return

                    elif key == ' ':  # 스페이스바
                        if current_index in selected:
                            selected.remove(current_index)
                        else:
                            selected.add(current_index)

                    elif key == '\r':  # Enter
                        break

            finally:
                termios.tcsetattr(sys.stdin, termios.TCSADRAIN, old_settings)

    except KeyboardInterrupt:
        print("\n\n작업이 중단되었습니다.")
        return

    if selected:
        selected_items = [items[i] for i in sorted(selected)]
        result_text = '\n\n'.join(selected_items)

        if copy_to_clipboard(result_text):
            clear_screen()
            print("=" * 60)
            print("클립보드에 복사 완료!")
            print("=" * 60)
            print("\n선택된 내용:")
            print("-" * 60)
            for item in selected_items:
                print(item)
                print()
            print("-" * 60)
            print(f"\n총 {len(selected_items)}개의 항목이 복사되었습니다.")
        else:
            print("\n클립보드 복사 실패. 수동으로 복사해주세요:")
            print("-" * 60)
            print(result_text)
            print("-" * 60)
    else:
        print("\n선택된 항목이 없습니다.")

if __name__ == "__main__":
    main()
