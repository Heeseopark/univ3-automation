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

import pandas as pd
import sys

try:
    # 임시 파일이 있으면 그것을 사용, 없으면 스크립트 실행해서 다운로드
    import os

    temp_file = "행정파일_생일_temp.xlsx"
    if not os.path.exists(temp_file):
        print("임시 파일이 없습니다. 다운로드를 시도합니다...")
        # birthday_selector의 download 함수 사용
        import subprocess
        from datetime import datetime, timedelta
        import pytz
        import gspread
        from google.oauth2 import service_account
        import requests
        import urllib3

        SHEET_ID = get_google_sheet_id()
        SCOPES = [
            'https://www.googleapis.com/auth/spreadsheets.readonly',
            'https://www.googleapis.com/auth/drive.readonly'
        ]

        credentials_paths = [str(path) for path in get_google_credentials_candidates()]

        creds_file = None
        for path in credentials_paths:
            if os.path.exists(path):
                creds_file = path
                break

        if not creds_file:
            print("credentials.json을 찾을 수 없습니다.")
            sys.exit(1)

        creds = service_account.Credentials.from_service_account_file(
            creds_file, scopes=SCOPES
        )

        session = requests.Session()
        session.verify = False
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

        from google.auth.transport.requests import Request as GoogleRequest

        class CustomRequest(GoogleRequest):
            def __init__(self):
                super().__init__(session=session)

        creds.refresh(CustomRequest())
        access_token = creds.token

        url = f"https://www.googleapis.com/drive/v3/files/{SHEET_ID}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        headers = {"Authorization": f"Bearer {access_token}"}

        response = session.get(url, headers=headers, timeout=30, verify=False)
        response.raise_for_status()

        with open(temp_file, 'wb') as f:
            f.write(response.content)

        print(f"다운로드 완료: {temp_file}")

    # 파일 읽기 - header 없이 전체 읽기
    df_no_header = pd.read_excel(temp_file, sheet_name='라인업정보', header=None)

    print(f"전체 행 개수: {len(df_no_header)}")
    print(f"전체 컬럼 개수: {len(df_no_header.columns)}")
    print(f"\n처음 10행 전체 확인:")
    print("=" * 120)
    print(df_no_header.head(10))

    print("\n" + "=" * 120)
    print("6번째 행(인덱스 5) 확인 (이것이 헤더일 가능성):")
    print(df_no_header.iloc[5])

    # header=5로 읽기
    df = pd.read_excel(temp_file, sheet_name='라인업정보', header=5)

    print("\n" + "=" * 120)
    print(f"header=5로 읽었을 때 컬럼 개수: {len(df.columns)}")
    print(f"\n컬럼 목록 (설정된 생년월일 열: AM, AN, AO):")
    for i, col in enumerate(df.columns):
        print(f"{i:3d}: {col}")

    print("\n" + "=" * 120)
    print("특정 컬럼의 데이터 샘플:")
    if len(df.columns) > 7:
        print(f"\n7번 컬럼 ({df.columns[7]}) 샘플:")
        print(df.iloc[:5, 7])
    if len(df.columns) > 6:
        print(f"\n6번 컬럼 ({df.columns[6]}) 샘플:")
        print(df.iloc[:5, 6])
    if len(df.columns) > 5:
        print(f"\n5번 컬럼 ({df.columns[5]}) 샘플:")
        print(df.iloc[:5, 5])
    if len(df.columns) > 4:
        print(f"\n4번 컬럼 ({df.columns[4]}) 샘플:")
        print(df.iloc[:5, 4])

except Exception as e:
    print(f"오류 발생: {e}")
    import traceback
    traceback.print_exc()


