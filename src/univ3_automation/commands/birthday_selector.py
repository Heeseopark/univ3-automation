#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
행정파일에서 생일자 추출 시스템
구글 스프레드시트에서 생일 데이터를 가져와 주차별 생일자를 추출합니다.
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
import pandas as pd
import subprocess
from datetime import datetime, timedelta
import pytz
import gspread
from google.oauth2 import service_account
import requests
import urllib3
import re

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

# Google Sheets 설정 (행정파일 - 라인업정보 시트)
SHEET_ID = get_google_sheet_id()
SHEET_NAME = "라인업정보"

# Google API 인증 설정
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/drive.readonly'
]

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def copy_to_clipboard(text):
    """클립보드에 텍스트 복사"""
    try:
        subprocess.run(
            ['powershell', '-Command', f'Set-Clipboard -Value @"\n{text}\n"@'],
            check=True,
            shell=False
        )
        return True
    except:
        return False

def get_week_ranges():
    """현재부터 4주간의 주차 범위 반환 (일요일 시작)"""
    kst = pytz.timezone('Asia/Seoul')
    now_kst = datetime.now(kst)

    # 이번 주 일요일 찾기
    weekday = now_kst.weekday()
    if weekday == 6:  # 이미 일요일
        this_sunday = now_kst
    else:
        days_until_sunday = 6 - weekday
        this_sunday = now_kst + timedelta(days=days_until_sunday)

    # 4주치 범위 생성
    weeks = []
    for i in range(4):
        week_start = this_sunday + timedelta(weeks=i)
        week_end = week_start + timedelta(days=6)
        weeks.append({
            'start': week_start,
            'end': week_end,
            'label': f"{week_start.month}/{week_start.day}(주일) ~ {week_end.month}/{week_end.day}(토)"
        })

    return weeks

def download_google_sheet():
    """구글 스프레드시트를 엑셀 파일로 다운로드 (Service Account 인증)"""
    print("구글 스프레드시트 다운로드 중...")

    try:
        # Service Account 인증 파일 경로
        credentials_paths = [str(path) for path in get_google_credentials_candidates()]

        creds_file = None
        for path in credentials_paths:
            if os.path.exists(path):
                creds_file = path
                break

        if not creds_file:
            raise FileNotFoundError("credentials.json not found")

        # Service Account 인증
        creds = service_account.Credentials.from_service_account_file(
            creds_file, scopes=SCOPES
        )

        # SSL 검증 우회 설정
        session = requests.Session()
        session.verify = False
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

        # 커스텀 Request 객체 생성
        from google.auth.transport.requests import Request as GoogleRequest

        class CustomRequest(GoogleRequest):
            def __init__(self):
                super().__init__(session=session)

        creds.refresh(CustomRequest())
        access_token = creds.token

        # Drive API를 통해 Excel 형식으로 다운로드
        url = f"https://www.googleapis.com/drive/v3/files/{SHEET_ID}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        headers = {"Authorization": f"Bearer {access_token}"}

        response = session.get(url, headers=headers, timeout=30, verify=False)
        response.raise_for_status()

        # 임시 파일로 저장
        temp_file = "행정파일_생일_temp.xlsx"
        with open(temp_file, 'wb') as f:
            f.write(response.content)

        print("✓ 스프레드시트 다운로드 완료")
        return temp_file

    except Exception as e:
        print(f"✗ 다운로드 오류: {e}")
        raise

def extract_birthdays(excel_file, week_ranges, selected_weeks):
    """선택된 주차의 생일자 추출"""
    # 데이터 읽기 (라인업정보 시트, header는 0행)
    df = pd.read_excel(excel_file, sheet_name='라인업정보', header=0)

    # 컬럼 인덱스:
    # 3번 = 리더 (D열)
    # 4번 = 성별 (E열)
    # 5번 = 학년 (F열)
    # 6번 = 이름 (G열)
    # 36번 = 년
    # 37번 = 월
    # 38번 = 일

    columns = df.columns.tolist()

    # 컬럼이 충분한지 확인
    if len(columns) > 38:
        name_col = columns[6]     # 이름 (G열)
        gender_col = columns[4]   # 성별 (E열)
        grade_col = columns[5]    # 학년 (F열)
        birth_year_col = columns[36]   # 년
        birth_month_col = columns[37]  # 월
        birth_day_col = columns[38]    # 일

        # 필요한 데이터만 추출
        df_filtered = df[[name_col, gender_col, grade_col, birth_year_col, birth_month_col, birth_day_col]].copy()
        df_filtered.columns = ['이름', '성별', '학년', '년', '생월', '생일']

        # 결측값 제거
        df_filtered = df_filtered.dropna(subset=['이름', '성별', '학년', '생월', '생일'])

        # 생월, 생일을 정수로 변환
        df_filtered['생월'] = pd.to_numeric(df_filtered['생월'], errors='coerce')
        df_filtered['생일'] = pd.to_numeric(df_filtered['생일'], errors='coerce')
        df_filtered = df_filtered.dropna(subset=['생월', '생일'])
        df_filtered['생월'] = df_filtered['생월'].astype(int)
        df_filtered['생일'] = df_filtered['생일'].astype(int)

        # 학년을 정수로 변환
        df_filtered['학년'] = pd.to_numeric(df_filtered['학년'], errors='coerce')
        df_filtered = df_filtered.dropna(subset=['학년'])
        df_filtered['학년'] = df_filtered['학년'].astype(int)

        # 성별은 이미 "남" 또는 "여"로 되어 있음

        # 선택된 주차들의 생일자 추출
        all_birthdays = []

        for week_idx in selected_weeks:
            week = week_ranges[week_idx]
            week_start = week['start']
            week_end = week['end']

            # 해당 주차의 생일자 필터링
            week_birthdays = []

            for _, row in df_filtered.iterrows():
                birth_month = int(row['생월'])
                birth_day = int(row['생일'])

                # 올해 생일 날짜 계산
                current_year = week_start.year
                try:
                    # 올해 생일
                    birthday_this_year = datetime(current_year, birth_month, birth_day)
                    # 내년 생일 (12월-1월 경계 처리)
                    birthday_next_year = datetime(current_year + 1, birth_month, birth_day)

                    # KST로 변환
                    kst = pytz.timezone('Asia/Seoul')
                    birthday_this_year = kst.localize(birthday_this_year)
                    birthday_next_year = kst.localize(birthday_next_year)

                    # 주차 범위 내에 생일이 있는지 확인
                    if (week_start.date() <= birthday_this_year.date() <= week_end.date()) or \
                       (week_start.date() <= birthday_next_year.date() <= week_end.date()):
                        week_birthdays.append(row)

                except ValueError:
                    # 잘못된 날짜 (예: 2월 30일)는 무시
                    continue

            all_birthdays.extend(week_birthdays)

        # 중복 제거
        if all_birthdays:
            df_birthdays = pd.DataFrame(all_birthdays).drop_duplicates(subset=['이름'])

            # 날짜별로 그룹화하기 위해 생월, 생일로 정렬 후 학년 내림차순, 이름 오름차순
            df_birthdays = df_birthdays.sort_values(
                by=['생월', '생일', '학년', '이름'],
                ascending=[True, True, False, True]
            )

            # 결과 포맷팅: 날짜별로 그룹화
            from collections import defaultdict
            date_groups = defaultdict(list)

            for _, row in df_birthdays.iterrows():
                gender = row['성별']
                grade = int(row['학년'])
                name = str(row['이름']).strip()

                # 이름 뒤의 숫자와 알파벳 제거
                name = re.sub(r'[0-9a-zA-Z]+$', '', name).strip()

                month = int(row['생월'])
                day = int(row['생일'])

                date_key = (month, day)
                date_groups[date_key].append(f"{gender}{grade} {name}")

            # 날짜순으로 정렬하여 출력
            result_lines = []
            for (month, day) in sorted(date_groups.keys()):
                date_str = f"{month:02d}.{day:02d}"
                people = ', '.join(date_groups[(month, day)])
                result_lines.append(f"{date_str}. – {people}")

            # Windows 줄바꿈 문자 사용 (PowerPoint 호환성)
            return '\r\n'.join(result_lines)
        else:
            return "해당 주차에 생일자가 없습니다."
    else:
        return "행정파일의 컬럼 구조가 예상과 다릅니다."

def display_menu(weeks, selected, current_index):
    """주차 선택 메뉴 표시"""
    clear_screen()
    print("=" * 60)
    print(f"{CYAN}생일자 추출 - 주차 선택{RESET}")
    print("=" * 60)
    print("↑/↓: 이동, 스페이스: 선택/해제, Enter: 완료, ESC: 취소")
    print("=" * 60)
    print()

    for i, week in enumerate(weeks):
        # 체크박스 표시
        selected_mark = f"{MAGENTA}[✓]{RESET}" if i in selected else "[ ]"

        # 커서 표시
        cursor = "▶ " if i == current_index else "  "

        # 색상 적용
        if i in selected:
            print(f"{cursor}{selected_mark} {MAGENTA}{i+1}. {week['label']}{RESET}")
        elif i == current_index:
            print(f"{cursor}{selected_mark} {CYAN}{i+1}. {week['label']}{RESET}")
        else:
            print(f"{cursor}{selected_mark} {i+1}. {week['label']}")

    print()
    print("=" * 60)
    print(f"선택된 주차: {len(selected)}개")

def main():
    """메인 함수"""
    try:
        print("=" * 60)
        print("생일자 추출 시스템")
        print("=" * 60)

        # 1. 주차 범위 계산
        weeks = get_week_ranges()

        # 2. 주차 선택 UI
        selected = set()
        current_index = 0

        if os.name == 'nt':
            import msvcrt

            while True:
                display_menu(weeks, selected, current_index)

                key = msvcrt.getch()

                if key == b'\xe0' or key == b'\x00':
                    key = msvcrt.getch()
                    if key == b'H':  # 위쪽 화살표
                        current_index = max(0, current_index - 1)
                    elif key == b'P':  # 아래쪽 화살표
                        current_index = min(len(weeks) - 1, current_index + 1)

                elif key == b' ':  # 스페이스바
                    if current_index in selected:
                        selected.remove(current_index)
                    else:
                        selected.add(current_index)

                elif key == b'\r':  # Enter
                    if selected:
                        break
                    else:
                        print("\n최소 1개 이상의 주차를 선택해주세요.")
                        import time
                        time.sleep(1)

                elif key == b'\x1b':  # ESC
                    print("\n작업이 취소되었습니다.")
                    return 0

        else:
            # Unix/Linux/Mac
            import termios, tty

            old_settings = termios.tcgetattr(sys.stdin)
            try:
                tty.setraw(sys.stdin.fileno())

                while True:
                    display_menu(weeks, selected, current_index)

                    key = sys.stdin.read(1)

                    if key == '\x1b':
                        next_key = sys.stdin.read(2)
                        if next_key == '[A':  # 위쪽 화살표
                            current_index = max(0, current_index - 1)
                        elif next_key == '[B':  # 아래쪽 화살표
                            current_index = min(len(weeks) - 1, current_index + 1)
                        elif next_key == '':
                            print("\n작업이 취소되었습니다.")
                            return 0

                    elif key == ' ':  # 스페이스바
                        if current_index in selected:
                            selected.remove(current_index)
                        else:
                            selected.add(current_index)

                    elif key == '\r':  # Enter
                        if selected:
                            break
                        else:
                            print("\n최소 1개 이상의 주차를 선택해주세요.")
                            import time
                            time.sleep(1)

            finally:
                termios.tcsetattr(sys.stdin, termios.TCSADRAIN, old_settings)

        # 3. 구글 스프레드시트 다운로드
        excel_file = download_google_sheet()

        # 4. 생일자 추출
        print("\n생일자 추출 중...")
        result_text = extract_birthdays(excel_file, weeks, selected)

        # 5. 클립보드에 복사
        if copy_to_clipboard(result_text):
            clear_screen()
            print("=" * 60)
            print("클립보드에 복사 완료!")
            print("=" * 60)
            print("\n선택된 주차:")
            for idx in sorted(selected):
                print(f"  - {weeks[idx]['label']}")
            print("\n생일자 명단:")
            print("-" * 60)
            print(result_text)
            print("-" * 60)
        else:
            print("\n클립보드 복사 실패. 수동으로 복사해주세요:")
            print("-" * 60)
            print(result_text)
            print("-" * 60)

        # 6. 임시 파일 정리
        if os.path.exists("행정파일_생일_temp.xlsx"):
            os.remove("행정파일_생일_temp.xlsx")

        print("\n" + "=" * 60)
        print("✓ 작업이 완료되었습니다!")
        print("=" * 60)

        return 0

    except Exception as e:
        print(f"\n✗ 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())

