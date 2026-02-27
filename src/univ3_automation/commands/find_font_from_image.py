"""
이미지에서 폰트를 찾는 스크립트
PyAutoGUI를 사용하여 Chrome에서 fontfont.app에 접속하고 가장 최근 스크린샷을 업로드합니다.
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

import pyautogui
import pyperclip
import time
import os
from pathlib import Path
import subprocess


def get_latest_screenshot():
    """
    스크린샷 폴더에서 가장 최근에 저장된 이미지 파일을 찾습니다.
    (이름순으로도 가장 마지막)
    """
    # 설정 파일 기반 스크린샷 경로
    screenshot_paths = [Path(path) for path in get_screenshot_dirs()]

    latest_image = None
    latest_time = 0
    latest_by_name = None

    for screenshot_dir in screenshot_paths:
        if not screenshot_dir.exists():
            continue

        # 이미지 파일 확장자
        image_extensions = ['.png', '.jpg', '.jpeg', '.bmp', '.gif']

        # 해당 폴더의 모든 이미지 파일 수집
        image_files = [
            f for f in screenshot_dir.glob('*')
            if f.suffix.lower() in image_extensions
        ]

        if not image_files:
            continue

        # 수정 시간 기준 최신 파일
        for file_path in image_files:
            mtime = file_path.stat().st_mtime
            if mtime > latest_time:
                latest_time = mtime
                latest_image = file_path

        # 이름순 기준 최신 파일 (알파벳/숫자 순으로 마지막)
        sorted_files = sorted(image_files, key=lambda x: x.name)
        if sorted_files:
            candidate = sorted_files[-1]
            if latest_by_name is None or candidate.name > latest_by_name.name:
                latest_by_name = candidate

    # 이름순으로 가장 마지막 파일 선택 (요구사항에 맞춰)
    if latest_by_name:
        print(f"선택된 이미지 (이름순 최신): {latest_by_name}")
        if latest_image and latest_image != latest_by_name:
            print(f"참고: 수정 시간 기준 최신: {latest_image}")
        return latest_by_name

    return latest_image


def find_font_from_screenshot():
    """
    Chrome을 열고 fontfont.app에서 스크린샷의 폰트를 찾습니다.
    """
    # 1. 가장 최근 스크린샷 찾기
    screenshot_path = get_latest_screenshot()

    if not screenshot_path:
        print("스크린샷을 찾을 수 없습니다.")
        return

    print(f"선택된 이미지: {screenshot_path}")

    # 2. Chrome 실행
    print("\nChrome을 실행합니다...")
    chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe"

    # Chrome이 다른 경로에 있을 수 있음
    alternative_paths = [
        "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
        str(Path.home() / "AppData/Local/Google/Chrome/Application/chrome.exe"),
    ]

    if not os.path.exists(chrome_path):
        for alt_path in alternative_paths:
            if os.path.exists(alt_path):
                chrome_path = alt_path
                break

    # URL을 클립보드에 복사
    url = "https://fontfont.app/"
    pyperclip.copy(url)

    try:
        subprocess.Popen([chrome_path, url])
        print(f"Chrome에서 {url} 를 열었습니다.")
    except Exception as e:
        print(f"Chrome 실행 오류: {e}")
        print("수동으로 Chrome을 열고 https://fontfont.app/ 에 접속해주세요.")
        input("준비가 되면 Enter를 누르세요...")

    # 3. 페이지 로딩 대기
    print("페이지 로딩 대기 중...")
    time.sleep(3)

    # 4. 'v' 입력
    print("'v' 키를 입력합니다...")
    pyautogui.press('v')
    time.sleep(0.5)

    # 5. 'l' 입력
    print("'l' 키를 입력합니다...")
    pyautogui.press('l')
    time.sleep(1)

    # 6. 이미지 첨부 (파일 경로 복사 붙여넣기)
    print("이미지 경로를 붙여넣습니다...")
    # 파일 경로를 클립보드에 복사
    pyperclip.copy(str(screenshot_path))
    time.sleep(0.3)
    # Ctrl+V로 붙여넣기
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)

    # Enter 키로 확인
    print("파일을 선택합니다...")
    pyautogui.press('enter')
    time.sleep(2)

    print("\n완료! 폰트 검색 결과를 확인하세요.")

    # 7. 첨부된 이미지 삭제 (선택사항)
    delete_choice = input("\n첨부된 이미지를 삭제하시겠습니까? (y/n): ")
    if delete_choice.lower() == 'y':
        try:
            os.remove(screenshot_path)
            print(f"이미지가 삭제되었습니다: {screenshot_path}")
        except Exception as e:
            print(f"이미지 삭제 오류: {e}")


if __name__ == "__main__":
    print("=" * 60)
    print("이미지에서 폰트 찾기 스크립트")
    print("=" * 60)
    print()

    find_font_from_screenshot()

