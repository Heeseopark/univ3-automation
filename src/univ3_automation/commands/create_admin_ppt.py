#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
행정광고 PPT 생성 스크립트
이번 주 일요일 날짜를 기준으로 새로운 행정광고 PPT를 생성합니다.
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
from datetime import datetime, timedelta
from pptx import Presentation
import pytz
import re
import win32com.client
import pythoncom

# ====================================
# 설정 - 마지막에 남겨둘 슬라이드 개수
# ====================================
SLIDES_TO_KEEP_AT_END = 3

def get_this_sunday():
    """현재 한국 시간 기준으로 이번 주 일요일 날짜를 반환"""
    # 한국 시간대 설정
    kst = pytz.timezone('Asia/Seoul')
    now_kst = datetime.now(kst)

    # 현재 날짜의 요일 (0=월요일, 6=일요일)
    weekday = now_kst.weekday()

    # 이번 주 일요일 계산
    if weekday == 6:  # 이미 일요일인 경우
        sunday = now_kst
    else:
        # 다음 일요일까지 남은 일수
        days_until_sunday = 6 - weekday
        sunday = now_kst + timedelta(days=days_until_sunday)

    return sunday

def get_last_sunday():
    """지난주 일요일 날짜를 반환"""
    this_sunday = get_this_sunday()
    last_sunday = this_sunday - timedelta(days=7)
    return last_sunday

def find_latest_ppt(admin_dir, reference_sunday, max_weeks_back=8):
    """
    행정광고 디렉토리에서 PPT 파일 찾기
    reference_sunday 기준으로 1주 전, 2주 전... 순서대로 검색하여 가장 가까운 이전 주 PPT 반환
    """
    print("\n이전 주 PPT 검색 중...")

    for weeks_back in range(1, max_weeks_back + 1):
        target_sunday = reference_sunday - timedelta(days=7 * weeks_back)
        target_date_str = target_sunday.strftime("%m%d")

        print(f"  [{weeks_back}주 전] {target_sunday.strftime('%Y-%m-%d')} ({target_date_str}) ", end="")

        # 해당 날짜의 PPT 찾기
        for file in os.listdir(admin_dir):
            if file.endswith('.pptx') and not file.startswith('~$'):
                match = re.match(rf'{target_date_str}\s*행정광고.*\.pptx', file)
                if match:
                    print(f"✓ 발견!")
                    print(f"\n→ 원본 PPT: {file}")
                    return file

        print("✗")

    # 최대 검색 범위에서도 찾지 못한 경우
    print(f"\n경고: {max_weeks_back}주 전까지 PPT를 찾을 수 없습니다.")
    print("가장 최근 PPT를 대신 사용합니다...\n")

    # 재귀 호출 (reference_sunday=None으로 호출하여 가장 최근 파일 반환)
    return find_latest_ppt(admin_dir, reference_sunday=None)

def open_ppt_for_review(ppt_path):
    """PPT 파일을 열어서 사용자가 확인할 수 있도록 함"""
    pythoncom.CoInitialize()
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1  # PowerPoint 창 표시
        powerpoint.WindowState = 1  # ppWindowNormal - 일반 크기로 표시
        abs_ppt_path = os.path.abspath(ppt_path)
        presentation = powerpoint.Presentations.Open(abs_ppt_path)
        return powerpoint, presentation
    except Exception as e:
        print(f"PPT 열기 중 오류 발생: {str(e)}")
        pythoncom.CoUninitialize()
        return None, None

def close_ppt(powerpoint, presentation):
    """열린 PPT 파일을 닫기"""
    try:
        if presentation:
            presentation.Close()
        if powerpoint:
            powerpoint.Quit()
    except:
        pass
    finally:
        pythoncom.CoUninitialize()

def create_admin_ppt():
    """행정광고 PPT 생성 메인 함수"""

    print("=" * 50)
    print("행정광고 PPT 생성 스크립트")
    print(f"마지막에 남겨둘 슬라이드 개수: {SLIDES_TO_KEEP_AT_END}개")
    print("=" * 50 + "\n")

    # 디렉토리 경로 설정
    parent_dir = get_base_dir()
    admin_dir = os.path.join(parent_dir, "행정광고")

    if not os.path.exists(admin_dir):
        print(f"오류: {admin_dir} 디렉토리를 찾을 수 없습니다.")
        return False

    # 날짜 계산
    this_sunday = get_this_sunday()
    last_sunday = get_last_sunday()

    this_sunday_str = this_sunday.strftime("%m%d")  # MMDD 형식
    last_sunday_str = last_sunday.strftime("%m%d")

    print(f"지난주 일요일: {last_sunday.strftime('%Y년 %m월 %d일')} ({last_sunday_str})")
    print(f"이번주 일요일: {this_sunday.strftime('%Y년 %m월 %d일')} ({this_sunday_str})")
    print()

    # 새 PPT 파일 경로
    new_ppt_name = f"{this_sunday_str} 행정광고.pptx"
    new_ppt_path = os.path.join(admin_dir, new_ppt_name)

    # 이미 파일이 존재하는 경우
    if os.path.exists(new_ppt_path):
        print(f"{this_sunday_str} 행정광고 ppt 파일이 이미 존재합니다.")
        return True

    # 파일이 없으면 복사 프로세스 진행
    try:
        print("행정광고 PPT 생성 프로세스를 시작합니다...\n")

        # 지난주 PPT 파일 찾기 (이번주 일요일 기준으로 1주씩 뒤로 검색)
        latest_ppt = find_latest_ppt(admin_dir, reference_sunday=this_sunday)

        if not latest_ppt:
            print("오류: 행정광고 PPT 파일을 찾을 수 없습니다.")
            return False

        source_ppt_path = os.path.join(admin_dir, latest_ppt)
        print(f"원본 PPT 파일: {latest_ppt}")

        # PPT 파일을 열어서 사용자가 확인하도록 함
        print(f"\nPPT 파일을 열고 있습니다...")
        print("PowerPoint 창에서 슬라이드를 확인하고 부서 광고 마지막 슬라이드 번호를 확인하세요.")
        print("(헤세드, 군지체, 리더스쿨 등 부서 광고가 끝나는 슬라이드 번호)\n")

        powerpoint, presentation = open_ppt_for_review(source_ppt_path)

        if not powerpoint or not presentation:
            print("PPT 파일을 열 수 없습니다.")
            return False

        # 사용자 입력 받기
        while True:
            try:
                last_univ_group_slide_str = input("부서 광고 마지막 슬라이드 번호를 입력하세요 (1부터 시작): ")
                last_univ_group_slide_number = int(last_univ_group_slide_str)
                if last_univ_group_slide_number < 1:
                    print("1 이상의 숫자를 입력해주세요.")
                    continue
                break
            except ValueError:
                print("올바른 숫자를 입력해주세요.")

        print(f"\n입력된 부서 광고 마지막 슬라이드 번호: {last_univ_group_slide_number}")
        print("PowerPoint를 닫고 작업을 진행합니다...\n")

        # PPT 닫기
        close_ppt(powerpoint, presentation)

        # PPT 파일 다시 열기 (처리용)
        print(f"PPT 파일을 처리 중...")
        prs = Presentation(source_ppt_path)
        total_slides = len(prs.slides)
        print(f"전체 슬라이드 수: {total_slides}")

        # 제거할 슬라이드 인덱스 계산
        slides_to_remove = []

        # last_univ_group_slide_number 이후부터 마지막 SLIDES_TO_KEEP_AT_END개 슬라이드 이전까지 제거
        if total_slides > last_univ_group_slide_number + SLIDES_TO_KEEP_AT_END:
            remove_start = last_univ_group_slide_number  # 0-indexed로는 last_univ_group_slide_number
            remove_end = total_slides - SLIDES_TO_KEEP_AT_END

            for i in range(remove_start, remove_end):
                slides_to_remove.append(i)

        print(f"\n유지할 슬라이드:")
        print(f"  - 1번 ~ {last_univ_group_slide_number}번 슬라이드")
        print(f"  - 마지막 {SLIDES_TO_KEEP_AT_END}개 슬라이드 ({', '.join([str(total_slides - i) + '번' for i in range(SLIDES_TO_KEEP_AT_END - 1, -1, -1)])})")

        if slides_to_remove:
            print(f"\n제거할 슬라이드 번호 (1-indexed): {[i+1 for i in slides_to_remove]}")
            print(f"제거할 슬라이드 수: {len(slides_to_remove)}")

            # 슬라이드 제거 (역순으로 제거해야 인덱스가 꼬이지 않음)
            for idx in reversed(sorted(slides_to_remove)):
                rId = prs.slides._sldIdLst[idx].rId
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[idx]
        else:
            print("\n제거할 슬라이드가 없습니다.")

        # 저장
        prs.save(new_ppt_path)
        print(f"\n새 PPT 파일 생성 완료: {new_ppt_name}")
        print("\n작업이 성공적으로 완료되었습니다!")
        print("\n※ 연합광고를 추가하려면 '연합광고 반영' 메뉴를 실행하세요.")
        return True

    except PermissionError:
        print("\n오류: PPT 파일에 접근할 수 없습니다.")
        print("PowerPoint에서 파일이 열려있다면 닫고 다시 시도해주세요.")
        return False
    except Exception as e:
        print(f"\n오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """메인 실행 함수"""
    # 실행
    success = create_admin_ppt()
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()

