#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
출석부 자동 생성 시스템
구글 스프레드시트에서 데이터를 가져와 월별 출석부를 생성합니다.
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
import openpyxl
import subprocess
from PyPDF2 import PdfMerger
import re
from datetime import datetime, timedelta
import calendar
import requests
import tempfile
from tqdm import tqdm
import pytz
import shutil
from pathlib import Path
import gspread
from google.oauth2 import service_account
from io import BytesIO
import ssl
import certifi

# 고을 번호별 이름
TOWN_NAMES = {
    1: '기쁨',
    2: '다정',
    3: '미쁨',
    4: '믿음',
    5: '사랑',
    6: '소망',
    7: '온유'
}

# Google Sheets 설정
SHEET_ID = get_google_sheet_id()
SHEET_NAME = "출석부"
EXPORT_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"

# Google API 인증 설정
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/drive.readonly'
]

def get_current_month_info():
    """이번 달의 정보를 반환"""
    kst = pytz.timezone('Asia/Seoul')
    now_kst = datetime.now(kst)
    return now_kst.year, now_kst.month

def get_next_month_info():
    """다음 달의 정보를 반환"""
    kst = pytz.timezone('Asia/Seoul')
    now_kst = datetime.now(kst)

    # 다음 달 계산
    if now_kst.month == 12:
        next_month = 1
        next_year = now_kst.year + 1
    else:
        next_month = now_kst.month + 1
        next_year = now_kst.year

    return next_year, next_month


def get_leader_dir() -> Path:
    return Path(get_dir("leader", "리더파일"))

def select_target_month():
    """사용자가 월을 선택하도록 프롬프트 표시"""
    current_year, current_month = get_current_month_info()
    next_year, next_month = get_next_month_info()

    print("\n출석부를 생성할 월을 선택하세요:")
    print(f"1. 이번 달 ({current_year}년 {current_month}월)")
    print(f"2. 다음 달 ({next_year}년 {next_month}월)")

    while True:
        try:
            choice = input("\n선택 (1 또는 2): ").strip()
            if choice == '1':
                print(f"\n선택: {current_year}년 {current_month}월")
                return current_year, current_month
            elif choice == '2':
                print(f"\n선택: {next_year}년 {next_month}월")
                return next_year, next_month
            else:
                print("잘못된 입력입니다. 1 또는 2를 입력해주세요.")
        except KeyboardInterrupt:
            print("\n\n작업이 취소되었습니다.")
            sys.exit(0)
        except Exception as e:
            print(f"입력 오류: {e}. 다시 시도해주세요.")

def get_sundays_in_month(year, month):
    """특정 월의 모든 일요일 날짜를 반환"""
    sundays = []
    cal = calendar.monthcalendar(year, month)

    for week in cal:
        if week[calendar.SUNDAY] != 0:
            sundays.append(week[calendar.SUNDAY])

    return sundays

def download_google_sheet():
    """구글 스프레드시트를 엑셀 파일로 다운로드 (Service Account 인증)"""
    print("구글 스프레드시트 다운로드 중...")

    try:
        # Service Account 인증 파일 경로 (코드 폴더 또는 상위 폴더에 위치)
        credentials_paths = [str(path) for path in get_google_credentials_candidates()]

        creds_file = None
        for path in credentials_paths:
            if os.path.exists(path):
                creds_file = path
                print(f"✓ 인증 파일 발견: {path}")
                break

        if not creds_file:
            print("✗ credentials.json 파일을 찾을 수 없습니다.")
            print("Google Cloud Console에서 Service Account 키를 다운로드하여")
            print("코드 폴더 또는 상위 폴더에 credentials.json으로 저장해주세요.")
            raise FileNotFoundError("credentials.json not found")

        # Service Account 인증
        creds = service_account.Credentials.from_service_account_file(
            creds_file, scopes=SCOPES
        )

        # gspread 클라이언트 생성
        client = gspread.authorize(creds)

        # Access Token 가져오기 - SSL 문제 해결을 위한 커스텀 Request 클래스 사용
        from google.auth.transport.requests import Request as GoogleRequest
        import urllib3
        from requests.adapters import HTTPAdapter

        # SSL 검증 우회 설정 (임시 해결책)
        session = requests.Session()
        session.verify = False
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

        # 커스텀 Request 객체 생성
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
        temp_file = "행정 파일_temp.xlsx"
        with open(temp_file, 'wb') as f:
            f.write(response.content)

        print("✓ 스프레드시트 다운로드 완료")
        return temp_file

    except FileNotFoundError:
        # credentials.json이 없는 경우 기존 파일 사용
        if os.path.exists("행정 파일.xlsx"):
            print("인증 파일이 없어 기존 파일을 사용합니다.")
            return "행정 파일.xlsx"
        else:
            raise
    except Exception as e:
        print(f"✗ 다운로드 오류: {e}")
        # 기존 파일이 있으면 사용
        if os.path.exists("행정 파일.xlsx"):
            print("기존 파일을 사용합니다.")
            return "행정 파일.xlsx"
        else:
            raise

def prepare_leader_file(year, month):
    """리더파일 준비 및 월 정보 업데이트"""
    sundays = get_sundays_in_month(year, month)
    num_sundays = len(sundays)

    # 현재 월 계산 (이전 달)
    current_month = month - 1 if month > 1 else 12

    print(f"\n{year}년 {month}월 일요일 정보:")
    print(f"- 일요일 개수: {num_sundays}개")
    print(f"- 날짜: {sundays}")

    leader_dir = get_leader_dir()

    # 리더파일 선택 (4주 또는 5주)
    if num_sundays == 4:
        template_file = leader_dir / "리더파일 (4주).xlsx"
    else:
        template_file = leader_dir / "리더파일 (5주).xlsx"

    if not template_file.exists():
        print(f"경고: {template_file}이 없습니다. 기본 템플릿을 사용합니다.")
        template_file = leader_dir / "리더파일 (4주).xlsx"

    # 작업용 리더파일 복사 (리더파일 폴더 내에서 작업)
    work_template = leader_dir / f"리더파일_{month}월.xlsx"
    shutil.copy(template_file, work_template)

    # 엑셀 파일 열기 및 월 정보 업데이트
    wb = openpyxl.load_workbook(str(work_template))
    ws = wb.active

    # B2 셀의 월 정보 수정 (고정된 형식으로 설정)
    b2_cell = ws['B2']
    b2_cell.value = f"대학부 예수사람 리더파일 - {month}월"

    # 일요일 날짜 입력 (6행) - 병합된 셀 그룹을 찾아서 처리
    from openpyxl.cell.cell import MergedCell

    # 6행에서 병합된 셀 그룹 찾기
    merged_groups = []
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_row == 6 and merged_range.max_row == 6:
            # 6행의 병합된 셀 그룹
            first_cell = ws.cell(merged_range.min_row, merged_range.min_col)
            # F열(6) 이후의 병합된 셀만 포함
            if merged_range.min_col >= 6:  # F열 = 6
                merged_groups.append({
                    'first_cell': first_cell.coordinate,
                    'col': merged_range.min_col
                })

    # 병합된 그룹을 열 번호로 정렬
    merged_groups.sort(key=lambda x: x['col'])

    # 날짜 입력할 셀 위치 결정
    date_cells = []
    if len(merged_groups) >= 4:
        # 병합된 그룹이 충분하면 각 그룹의 첫 번째 셀 사용
        for group in merged_groups[:min(num_sundays, len(merged_groups))]:
            date_cells.append(group['first_cell'])
    else:
        # 병합된 그룹이 부족하면 기본 위치 사용
        if num_sundays == 4:
            date_cells = ['F6', 'H6', 'J6', 'L6']
        else:
            date_cells = ['F6', 'H6', 'J6', 'L6', 'N6']

    # 날짜 입력
    for i, sunday_date in enumerate(sundays):
        if i < len(date_cells):
            cell_addr = date_cells[i]
            cell = ws[cell_addr]

            # 병합된 셀이면 건너뛰기 (첫 번째 셀만 처리)
            if isinstance(cell, MergedCell):
                for merged_range in ws.merged_cells.ranges:
                    if cell.coordinate in merged_range:
                        first_cell = ws.cell(merged_range.min_row, merged_range.min_col)
                        first_cell.value = f"{month}월 {sunday_date}일"
                        break
            else:
                cell.value = f"{month}월 {sunday_date}일"

    wb.save(str(work_template))
    wb.close()

    print(f"✓ 리더파일 준비 완료: {work_template}")

    # 리더파일 자동으로 열기
    print("\n" + "=" * 60)
    print("리더파일 템플릿을 열어서 검토 및 수정하세요.")
    print("=" * 60)
    print(f"\n리더파일을 여는 중: {work_template}")

    try:
        # Windows에서 기본 프로그램으로 파일 열기
        if sys.platform == 'win32':
            os.startfile(os.path.abspath(work_template))
        elif sys.platform == 'darwin':  # macOS
            subprocess.call(['open', str(work_template)])
        else:  # Linux
            subprocess.call(['xdg-open', str(work_template)])

        print("\n✓ 파일이 열렸습니다.")
        print("수정이 필요하면 수정하고 파일을 저장한 후 닫아주세요.")
        input("\n작업이 완료되면 Enter 키를 눌러 계속 진행하세요...")
        print("\n✓ 작업을 계속합니다.")

    except Exception as e:
        print(f"\n파일 열기 실패: {e}")
        print("수동으로 파일을 열어 확인 및 수정해주세요.")
        input("확인이 완료되면 Enter 키를 눌러 계속 진행하세요...")

    # 사용자 수정사항 반영을 위해 파일 다시 로드
    print("\n사용자 수정사항을 반영 중...")

    return str(work_template), num_sundays

def process_attendance_data(excel_file, template_file, year, month):
    """출석부 데이터 처리 및 PDF 생성"""

    # 데이터 읽기
    print("\n출석부 데이터 처리 중...")
    df = pd.read_excel(excel_file, sheet_name='출석부', header=5, dtype={'KEY': str, '조원': str})

    # 데이터 전처리
    df['group_key'] = df['KEY'].str[:2]
    valid_town_numbers = [str(key) for key in TOWN_NAMES.keys()]

    filtered_df = df[
        (df['KEY'].str.len() == 4) &
        (df['group_key'].str[1] != '0') &
        df['조원'].str.strip().ne('') &
        df['group_key'].str[0].isin(valid_town_numbers)
    ]

    # 그룹별 처리
    grouped = filtered_df.groupby('group_key')

    leader_dir = get_leader_dir()
    temp_dir = leader_dir / "temp"

    # temp 디렉토리 생성 (리더파일 폴더 내에)
    temp_dir.mkdir(parents=True, exist_ok=True)

    excel_files = []
    group_info = []  # progress bar용 정보 저장

    # 먼저 그룹 정보 수집
    for group, data in grouped:
        group_str = str(group)
        town_number = int(group_str[0])
        group_number = int(group_str[1])
        town_name = TOWN_NAMES.get(town_number)

        data = data.dropna(subset=['조원'])
        if not data.empty:
            first_member = data.iloc[0]['조원'].split()[-1]
            first_member = re.sub(r'[0-9]{1,2}[a-zA-Z]?$', '', first_member)
            group_info.append({
                'group': group,
                'town_name': town_name,
                'group_number': group_number,
                'leader': first_member,
                'data': data
            })

    # Progress bar로 처리
    print("\n출석부 생성 중...")
    with tqdm(total=len(group_info), desc="처리 진행", unit="고을") as pbar:
        for info in group_info:
            # Progress bar 설명 업데이트
            pbar.set_description(f"{info['town_name']}-{info['leader']} GBS")

            wb = openpyxl.load_workbook(template_file)
            ws = wb.active

            # 고을 정보 입력
            ws['B4'] = f"고을 이름: {info['town_name']}"
            ws['E4'] = f"GBS 리더명: {info['leader']}"

            # 조원 정보 입력
            for index, row in enumerate(info['data'].itertuples(), start=11):
                if pd.notna(row.조원):
                    parts = row.조원.split()
                    if len(parts) >= 2:
                        gender_grade = parts[0]
                        name = parts[1]

                        # 이름 정제
                        name = re.sub(r'[0-9]{1,2}[a-zA-Z]?$', '', name)
                        name = re.sub(r'[a-zA-Z]$', '', name)

                        gender = gender_grade[0]
                        grade = gender_grade[1:]

                        ws[f'C{index}'] = grade
                        ws[f'D{index}'] = gender
                        ws[f'E{index}'] = name

            # 파일 저장 (리더파일/temp 폴더에)
            temp_excel_filename = temp_dir / f"{info['town_name']}_{info['group_number']}.xlsx"
            wb.save(str(temp_excel_filename))
            excel_files.append(str(temp_excel_filename))

            pbar.update(1)

    print(f"✓ {len(excel_files)}개 출석부 생성 완료")
    return excel_files

def convert_to_pdf(excel_files):
    """엑셀 파일을 PDF로 변환 (convert.py 로직 참고)"""
    print("\nPDF 변환 중...")

    # LibreOffice 경로 설정
    libreoffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"
    if not os.path.exists(libreoffice_path):
        libreoffice_path = r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
    if not os.path.exists(libreoffice_path):
        print("경고: LibreOffice를 찾을 수 없습니다. PDF 변환을 건너뜁니다.")
        return []

    pdf_files = []

    with tqdm(total=len(excel_files), desc="PDF 변환", unit="파일") as pbar:
        for excel_file in excel_files:
            pdf_file = excel_file.replace('.xlsx', '.pdf')

            try:
                subprocess.call([
                    libreoffice_path,
                    '--headless',
                    '--convert-to', 'pdf',
                    '--outdir', os.path.dirname(excel_file),
                    excel_file
                ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

                pdf_files.append(pdf_file)
            except Exception as e:
                print(f"변환 실패: {excel_file} - {e}")

            pbar.update(1)

    print(f"✓ {len(pdf_files)}개 PDF 생성 완료")
    return pdf_files

def merge_pdfs(pdf_files, year, month):
    """PDF 파일들을 하나로 합치기"""
    if not pdf_files:
        print("합칠 PDF 파일이 없습니다.")
        return None

    print("\nPDF 병합 중...")

    merger = PdfMerger()

    # 파일명으로 정렬하여 순서대로 병합
    pdf_files.sort()

    for pdf_file in pdf_files:
        if os.path.exists(pdf_file):
            merger.append(pdf_file)

    # 최종 파일명 생성 (리더파일 폴더에 저장)
    output_filename = str(get_leader_dir() / f"{month}월 출석 파일.pdf")
    merger.write(output_filename)
    merger.close()

    print(f"✓ 최종 파일 생성: {output_filename}")
    return output_filename

def cleanup_temp_files(excel_files, pdf_files):
    """임시 파일 정리"""
    print("\n임시 파일 정리 중...")

    for file in excel_files + pdf_files:
        try:
            if os.path.exists(file):
                os.remove(file)
        except Exception as e:
            print(f"파일 삭제 실패: {file} - {e}")

    # 리더파일 폴더의 temp 디렉토리가 비어있으면 삭제
    temp_dir = get_leader_dir() / "temp"
    if temp_dir.exists() and not any(temp_dir.iterdir()):
        temp_dir.rmdir()

    print("✓ 임시 파일 정리 완료")

def main():
    """메인 함수"""
    try:
        print("=" * 60)
        print("출석부 자동 생성 시스템")
        print("=" * 60)

        # 사용자가 월 선택
        year, month = select_target_month()
        print(f"\n대상: {year}년 {month}월")

        # 1. 구글 스프레드시트 다운로드
        excel_file = download_google_sheet()

        # 2. 리더파일 준비
        template_file, num_weeks = prepare_leader_file(year, month)

        # 3. 출석부 데이터 처리
        excel_files = process_attendance_data(excel_file, template_file, year, month)

        # 4. PDF 변환
        pdf_files = convert_to_pdf(excel_files)

        # 5. PDF 병합
        if pdf_files:
            final_pdf = merge_pdfs(pdf_files, year, month)

            # 6. 임시 파일 정리
            cleanup_temp_files(excel_files, pdf_files)

            # 임시 다운로드 파일 정리
            if os.path.exists("행정 파일_temp.xlsx"):
                os.remove("행정 파일_temp.xlsx")

            # 리더파일 정리 (출석부 생성 완료 후)
            work_template = get_leader_dir() / f"리더파일_{month}월.xlsx"
            if work_template.exists():
                os.remove(str(work_template))
                print(f"✓ 임시 리더파일 삭제: 리더파일_{month}월.xlsx")

        print("\n" + "=" * 60)
        print("✓ 모든 작업이 완료되었습니다!")
        print("=" * 60)

        return 0

    except Exception as e:
        print(f"\n✗ 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
