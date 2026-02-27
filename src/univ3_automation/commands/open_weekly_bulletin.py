#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
주보 PPT 파일 열기
주보 폴더에 없으면 카카오톡 받은 파일에서 찾아서 이동 후 열기
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
import subprocess
import shutil
import re
from datetime import datetime, timedelta
import pytz
from pathlib import Path

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

def find_bulletin_file(bulletin_dir, mmdd):
    """주보 폴더에서 해당 날짜의 주보 파일 찾기"""
    files = list(Path(bulletin_dir).glob("*.pptx"))

    # 파일명에 주보가 포함되고 날짜가 맞는 파일 찾기
    for file in files:
        if "주보" in file.name and mmdd in file.name:
            return file

    return None

def find_bulletin_in_kakao(kakao_dir, mmdd):
    """카카오톡 다운로드 폴더에서 주보 파일 찾기"""
    if not os.path.exists(kakao_dir):
        return None

    files = list(Path(kakao_dir).glob("*.pptx"))

    # 파일명에 주보가 포함되고 날짜가 맞는 파일 찾기
    for file in files:
        if "주보" in file.name and mmdd in file.name:
            return file

    # 가장 최근에 수정된 주보 파일 찾기
    bulletin_files = [f for f in files if "주보" in f.name]
    if bulletin_files:
        bulletin_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        return bulletin_files[0]

    return None

def move_bulletin_from_kakao(source_file, bulletin_dir, mmdd):
    """카카오톡 폴더에서 주보 폴더로 파일 이동"""
    # 호수 추출 (예: "1579호" -> 1579)
    match = re.search(r'(\d+)호', source_file.name)
    issue_number = match.group(1) if match else "0000"

    # 타겟 파일명 생성
    target_filename = f"{issue_number}호 {mmdd} 주보.pptx"
    target_path = Path(bulletin_dir) / target_filename

    # 타겟 폴더가 없으면 생성
    Path(bulletin_dir).mkdir(parents=True, exist_ok=True)

    # 파일 이동 (복사 후 삭제)
    shutil.copy2(source_file, target_path)
    source_file.unlink()  # 원본 삭제

    print(f"✓ 주보 파일을 이동했습니다:")
    print(f"  {source_file.name} → {target_filename}")

    return target_path

def main():
    """메인 함수"""
    try:
        # 이번주 일요일 날짜 계산
        sunday = get_this_sunday()
        mmdd = sunday.strftime("%m%d")

        # 주보 디렉토리 경로
        base_dir = get_base_dir()
        bulletin_dir = os.path.join(base_dir, "주보")

        print(f"\n{mmdd} 주보 파일을 찾는 중...")

        # 1. 주보 폴더에서 찾기
        bulletin_file = find_bulletin_file(bulletin_dir, mmdd)

        if not bulletin_file:
            # 2. 카카오톡 받은 파일에서 찾기
            kakao_dir = get_kakao_downloads_dir()
            print(f"주보 폴더에 없습니다. 카카오톡 받은 파일에서 검색 중...")

            kakao_file = find_bulletin_in_kakao(kakao_dir, mmdd)

            if not kakao_file:
                print(f"\n❌ {mmdd} 날짜의 주보 파일을 찾을 수 없습니다.")
                print(f"   주보 폴더: {bulletin_dir}")
                print(f"   카카오톡 폴더: {kakao_dir}")
                print("\n주보 파일을 확인해주세요.")
                return

            # 파일 이동
            bulletin_file = move_bulletin_from_kakao(kakao_file, bulletin_dir, mmdd)
        else:
            print(f"✓ 찾은 파일: {bulletin_file.name}")

        # 파일 열기
        print(f"\n{bulletin_file.name} 파일을 엽니다...")

        if os.name == 'nt':  # Windows
            os.startfile(bulletin_file)
        elif sys.platform == 'darwin':  # macOS
            subprocess.run(['open', str(bulletin_file)])
        else:  # Linux
            subprocess.run(['xdg-open', str(bulletin_file)])

        print("파일이 열렸습니다!")

    except Exception as e:
        print(f"\n오류 발생: {str(e)}")
        return

if __name__ == "__main__":
    main()

