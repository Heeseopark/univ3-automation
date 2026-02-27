#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
카카오톡 플러스친구 PPT 파일 열기
카톡플친 폴더에서 가장 최근 파일 자동으로 열기
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

def find_kakao_plus_friend_file(kakao_dir, mmdd):
    """카톡플친 폴더에서 해당 날짜의 PPT 파일 찾기"""
    if not os.path.exists(kakao_dir):
        return None

    files = list(Path(kakao_dir).glob("*.pptx"))

    # 파일명에 날짜가 맞는 파일 찾기
    for file in files:
        if mmdd in file.name:
            return file

    # 가장 최근에 수정된 PPT 파일 찾기
    if files:
        files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        return files[0]

    return None

def main():
    """메인 함수"""
    try:
        # 이번주 일요일 날짜 계산
        sunday = get_this_sunday()
        mmdd = sunday.strftime("%m%d")

        # 카톡플친 디렉토리 경로
        base_dir = get_base_dir()
        kakao_dir = os.path.join(base_dir, "카톡플친")

        print(f"\n{mmdd} 카톡플친 PPT 파일을 찾는 중...")

        # 카톡플친 폴더에서 찾기
        kakao_file = find_kakao_plus_friend_file(kakao_dir, mmdd)

        if not kakao_file:
            print(f"\n❌ {mmdd} 날짜의 카톡플친 PPT 파일을 찾을 수 없습니다.")
            print(f"   카톡플친 폴더: {kakao_dir}")
            print("\n카톡플친 PPT 파일을 확인해주세요.")
            return

        print(f"✓ 찾은 파일: {kakao_file.name}")

        # 파일 열기
        print(f"\n{kakao_file.name} 파일을 엽니다...")

        if os.name == 'nt':  # Windows
            os.startfile(kakao_file)
        elif sys.platform == 'darwin':  # macOS
            subprocess.run(['open', str(kakao_file)])
        else:  # Linux
            subprocess.run(['xdg-open', str(kakao_file)])

        print("파일이 열렸습니다!")

    except Exception as e:
        print(f"\n오류 발생: {str(e)}")
        return

if __name__ == "__main__":
    main()

