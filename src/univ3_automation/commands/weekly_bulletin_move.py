#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
주보 파일 이동 스크립트
카카오톡에서 다운로드한 주보 파일을 지정된 위치로 이동합니다.
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
import shutil
from datetime import datetime, timedelta
import pytz
from pathlib import Path
import re

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

def find_bulletin_file(source_dir, mmdd):
    """카카오톡 다운로드 폴더에서 주보 파일 찾기"""
    files = list(Path(source_dir).glob("*.pptx"))

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

def move_bulletin():
    """주보 파일을 지정된 위치로 이동"""
    # 날짜 계산
    sunday = get_this_sunday()
    mmdd = sunday.strftime("%m%d")

    # 소스 디렉토리 (카카오톡 다운로드 폴더)
    source_dir = get_kakao_downloads_dir()

    # 타겟 디렉토리 (주보 폴더)
    target_dir = get_dir("bulletin", "주보")

    print("=" * 60)
    print(f"주보 파일 이동 - {sunday.strftime('%Y년 %m월 %d일')}")
    print("=" * 60)
    print()

    # 소스 파일 찾기
    print(f"소스 폴더에서 주보 파일 검색 중...")
    print(f"폴더: {source_dir}")

    source_file = find_bulletin_file(source_dir, mmdd)

    if not source_file:
        print("\n❌ 주보 파일을 찾을 수 없습니다.")
        print(f"   {mmdd} 날짜가 포함된 주보 파일을 확인해주세요.")
        return False

    print(f"✓ 찾은 파일: {source_file.name}")

    # 호수 추출 (예: "1579호" -> 1579)
    match = re.search(r'(\d+)호', source_file.name)
    issue_number = match.group(1) if match else "0000"

    # 타겟 파일명 생성
    target_filename = f"{issue_number}호 {mmdd} 주보.pptx"
    target_path = Path(target_dir) / target_filename

    print(f"\n타겟 위치: {target_path}")

    # 파일 이동
    try:
        # 타겟 폴더가 없으면 생성
        Path(target_dir).mkdir(parents=True, exist_ok=True)

        # 파일 이동 (복사 후 삭제)
        shutil.copy2(source_file, target_path)
        source_file.unlink()  # 원본 삭제

        print("\n✓ 주보 파일이 성공적으로 이동되었습니다!")
        print(f"  {source_file.name}")
        print(f"  → {target_filename}")
        return True

    except Exception as e:
        print(f"\n❌ 파일 이동 중 오류 발생: {str(e)}")
        return False

def main():
    """메인 함수"""
    success = move_bulletin()

    print("\n" + "=" * 60)
    if success:
        print("작업 완료!")
    else:
        print("작업 실패")
    print("=" * 60)

if __name__ == "__main__":
    main()

