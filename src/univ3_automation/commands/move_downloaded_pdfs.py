#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
다운로드 폴더의 최근 PDF 파일들을 25-2 루트 디렉토리로 이동
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
import shutil
from pathlib import Path
from datetime import datetime

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
        elif key == b' ':  # Space
            return 'SPACE'
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
            elif key == ' ':  # Space
                return 'SPACE'
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

def select_pdfs(pdf_files):
    """이동할 PDF 파일 선택 - 화살표 키로 이동, 스페이스로 선택"""
    selected = [False] * len(pdf_files)
    current_index = 0

    while True:
        clear_screen()
        print("=" * 80)
        print("이동할 파일을 선택하세요 (스페이스: 선택/해제, Enter: 완료, Q: 취소)")
        print("=" * 80)
        print()

        for i, pdf_file in enumerate(pdf_files):
            file_size = pdf_file.stat().st_size / 1024  # KB
            file_time = datetime.fromtimestamp(pdf_file.stat().st_mtime)

            checkbox = "[✓]" if selected[i] else "[ ]"
            prefix = "▶ " if i == current_index else "  "

            print(f"{prefix}{checkbox} {i+1}. {pdf_file.name}")
            print(f"     크기: {file_size:.1f} KB | 수정 시간: {file_time.strftime('%Y-%m-%d %H:%M:%S')}")
            print()

        print("=" * 80)
        selected_count = sum(selected)
        print(f"선택됨: {selected_count}개")

        key = get_key()

        if key == 'UP':
            current_index = (current_index - 1) % len(pdf_files)
        elif key == 'DOWN':
            current_index = (current_index + 1) % len(pdf_files)
        elif key == 'SPACE':
            selected[current_index] = not selected[current_index]
        elif key == 'ENTER':
            return [i for i, s in enumerate(selected) if s]
        elif key == 'Q':
            return []

def move_pdfs(pdf_files, selected_indices, target_dir):
    """선택된 PDF 파일들을 타겟 디렉토리로 이동"""
    if not selected_indices:
        print("\n선택된 파일이 없습니다.")
        return False, []

    print("\n" + "=" * 80)
    print("파일 이동 중...")
    print("=" * 80)

    # 타겟 디렉토리가 없으면 생성
    os.makedirs(target_dir, exist_ok=True)

    moved_count = 0
    moved_files = []

    for idx in selected_indices:
        source_file = pdf_files[idx]
        target_file = Path(target_dir) / source_file.name

        # 같은 이름의 파일이 이미 있으면 숫자 추가
        counter = 1
        original_name = source_file.stem
        while target_file.exists():
            target_file = Path(target_dir) / f"{original_name}_{counter}.pdf"
            counter += 1

        try:
            # 파일 이동 (복사 후 삭제)
            shutil.copy2(source_file, target_file)
            source_file.unlink()  # 원본 삭제

            print(f"✓ {source_file.name}")
            if counter > 1:
                print(f"  → {target_file.name} (이름 변경됨)")

            moved_count += 1
            moved_files.append(target_file)
        except Exception as e:
            print(f"✗ {source_file.name}: {str(e)}")

    print("\n" + "=" * 80)
    print(f"{moved_count}개 파일 이동 완료!")
    return moved_count > 0, moved_files

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
    print("다운로드 PDF 파일 이동")
    print("=" * 80)

    # 소스 폴더 선택
    download_dir = select_source_folder()

    # 타겟 디렉토리 (25-2 루트)
    base_dir = get_base_dir()
    target_dir = base_dir

    print(f"\n소스 폴더: {download_dir}")
    print(f"이동 대상 폴더: {target_dir}")

    # 최근 PDF 파일 찾기
    pdf_files = get_recent_pdfs(download_dir, count=5)

    if not pdf_files:
        print("\n다운로드 폴더에 PDF 파일이 없습니다.")
        return

    # PDF 목록 표시
    display_pdf_list(pdf_files)

    # 파일 선택
    selected_indices = select_pdfs(pdf_files)

    if not selected_indices:
        print("\n작업을 취소했습니다.")
        return

    # 파일 이동
    success, moved_files = move_pdfs(pdf_files, selected_indices, target_dir)

    if success:
        print(f"\n파일이 성공적으로 이동되었습니다!")
        print(f"위치: {target_dir}")

        # 파일 열기 여부 확인
        print("\n이동된 파일을 열겠습니까? (Y/N): ", end="")
        choice = input().strip().upper()

        if choice == 'Y':
            import subprocess
            for file in moved_files:
                try:
                    if os.name == 'nt':  # Windows
                        os.startfile(file)
                    elif sys.platform == 'darwin':  # macOS
                        subprocess.run(['open', str(file)])
                    else:  # Linux
                        subprocess.run(['xdg-open', str(file)])
                    print(f"✓ {file.name} 열기")
                except Exception as e:
                    print(f"✗ {file.name} 열기 실패: {str(e)}")

if __name__ == "__main__":
    main()

