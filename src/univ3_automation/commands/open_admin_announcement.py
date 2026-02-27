#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
행정광고 PPT 파일 열기
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

def main():
    """메인 함수"""
    try:
        # 이번주 일요일 날짜 계산
        sunday = get_this_sunday()
        mmdd = sunday.strftime("%m%d")

        # 행정광고 디렉토리 경로
        base_dir = get_base_dir()
        admin_dir = os.path.join(base_dir, "행정광고")

        # 행정광고 파일 경로
        ppt_path = os.path.join(admin_dir, f"{mmdd} 행정광고.pptx")

        # 파일 존재 확인
        if not os.path.exists(ppt_path):
            print(f"\n오류: {mmdd} 행정광고.pptx 파일을 찾을 수 없습니다.")
            print(f"경로: {ppt_path}")
            print("\n먼저 '행정광고 PPT 생성' 기능을 실행해주세요.")
            return

        # 파일 열기
        print(f"\n{mmdd} 행정광고.pptx 파일을 엽니다...")

        if os.name == 'nt':  # Windows
            os.startfile(ppt_path)
        elif sys.platform == 'darwin':  # macOS
            subprocess.run(['open', ppt_path])
        else:  # Linux
            subprocess.run(['xdg-open', ppt_path])

        print("파일이 열렸습니다!")

    except Exception as e:
        print(f"\n오류 발생: {str(e)}")
        return

if __name__ == "__main__":
    main()

