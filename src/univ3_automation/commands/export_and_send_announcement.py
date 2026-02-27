#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
행정광고 PPT 내보내기 및 이메일 전송 스크립트
이번 주 일요일 행정광고를 PDF/PNG로 변환하고 Gmail로 전송합니다.
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
from datetime import datetime, timedelta
import pytz
import win32com.client
import pythoncom
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email import encoders
import getpass
from pathlib import Path

# 색상 코드 정의
try:
    import colorama
    colorama.init()
    MAGENTA = '\033[95m'
    CYAN = '\033[96m'
    WHITE = '\033[97m'
    RESET = '\033[0m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
except ImportError:
    MAGENTA = CYAN = WHITE = RESET = GREEN = YELLOW = ''

# ====================================
# 설정
# ====================================
RECIPIENTS = [
    "heeseopark99@gmail.com",
    "univ3media@naver.com"
    # 필요시 추가 수신자를 여기에 추가
]

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def select_recipients(recipients_list):
    """수신자 선택 화면"""
    selected_recipients = set(range(len(recipients_list)))  # 기본적으로 모두 선택
    current_index = 0

    try:
        if os.name == 'nt':
            import msvcrt

            while True:
                clear_screen()
                print("=" * 60)
                print(f"{CYAN}이메일 수신자 선택{RESET}")
                print("=" * 60)
                print("↑/↓: 이동, 스페이스: 선택/해제, Enter: 확인")
                print("=" * 60)
                print()

                for i, recipient in enumerate(recipients_list):
                    selected_mark = f"{MAGENTA}[✓]{RESET}" if i in selected_recipients else "[ ]"
                    cursor = "▶ " if i == current_index else "  "

                    if i in selected_recipients:
                        print(f"{cursor}{selected_mark} {MAGENTA}{recipient}{RESET}")
                    elif i == current_index:
                        print(f"{cursor}{selected_mark} {CYAN}{recipient}{RESET}")
                    else:
                        print(f"{cursor}{selected_mark} {recipient}")

                print()
                print("=" * 60)
                print(f"선택된 수신자: {len(selected_recipients)}명")

                key = msvcrt.getch()

                if key == b'\xe0' or key == b'\x00':
                    key = msvcrt.getch()
                    if key == b'H':  # 위쪽 화살표
                        current_index = max(0, current_index - 1)
                    elif key == b'P':  # 아래쪽 화살표
                        current_index = min(len(recipients_list) - 1, current_index + 1)

                elif key == b' ':  # 스페이스바
                    if current_index in selected_recipients:
                        selected_recipients.remove(current_index)
                    else:
                        selected_recipients.add(current_index)

                elif key == b'\r':  # Enter
                    break

                elif key == b'\x1b':  # ESC
                    return None

    except KeyboardInterrupt:
        return None

    if selected_recipients:
        return [recipients_list[i] for i in sorted(selected_recipients)]
    return None

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

def export_ppt_to_pdf_and_png(ppt_path, output_dir):
    """PPT를 PDF와 PNG로 변환"""
    print(f"\nPPT 파일 변환 시작: {os.path.basename(ppt_path)}")

    pythoncom.CoInitialize()

    try:
        # PowerPoint 애플리케이션 시작
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1  # PowerPoint는 반드시 Visible이어야 함
        powerpoint.WindowState = 2  # ppWindowMinimized - 최소화 상태로 실행

        # PPT 파일 열기
        abs_ppt_path = os.path.abspath(ppt_path)
        presentation = powerpoint.Presentations.Open(abs_ppt_path)

        # PDF 경로 설정
        base_name = os.path.basename(ppt_path).replace('.pptx', '')
        pdf_path = os.path.join(output_dir, f"{base_name} PDF.pdf")

        # PDF로 저장
        print(f"PDF 생성 중...")
        presentation.SaveAs(pdf_path, 32)  # 32 = ppSaveAsPDF
        print(f"PDF 생성 완료: {os.path.basename(pdf_path)}")

        # PNG 디렉토리 생성
        png_dir = os.path.join(output_dir, f"{base_name} PNG")
        if os.path.exists(png_dir):
            shutil.rmtree(png_dir)
        os.makedirs(png_dir)

        # PNG로 저장 (전체 슬라이드 추출)
        # PDF처럼 원본 품질 그대로 추출 (크기 지정 없이 Export)
        total_slides = presentation.Slides.Count
        print(f"\nPNG 추출 중 (1번째 ~ {total_slides}번째 슬라이드)...")

        slide_num = 1  # 파일명용 번호
        for i in range(1, total_slides + 1):  # PowerPoint는 1-based, 첫 번째 슬라이드부터 마지막까지 전체
            slide = presentation.Slides(i)  # COM 객체는 괄호 사용
            png_file = os.path.join(png_dir, f"슬라이드{slide_num:02d}.png")
            # 크기 지정 없이 Export하면 원본 품질 그대로 추출됨
            slide.Export(png_file, "PNG")
            print(f"  - 슬라이드 {i} 추출 완료")
            slide_num += 1

        print(f"PNG 추출 완료: {png_dir}")

        # PowerPoint 종료
        presentation.Close()
        powerpoint.Quit()

        return pdf_path, png_dir

    except Exception as e:
        print(f"변환 중 오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return None, None

    finally:
        pythoncom.CoUninitialize()

def send_email_with_attachments(pdf_path, sunday_str, recipients):
    """PDF를 Gmail로 전송"""
    print(f"\n{CYAN}이메일 전송 준비...{RESET}")

    # Gmail 설정
    sender_email = recipients[0]  # 첫 번째 수신자를 발신자로 사용

    # 이메일 구성
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ', '.join(recipients)
    msg['Subject'] = f"{sunday_str} 행정광고"

    # PDF 첨부
    if pdf_path and os.path.exists(pdf_path):
        with open(pdf_path, "rb") as f:
            part = MIMEApplication(f.read(), _subtype='pdf')
            part.add_header(
                'Content-Disposition',
                'attachment',
                filename=os.path.basename(pdf_path)
            )
            msg.attach(part)
        print(f"PDF 첨부 완료: {os.path.basename(pdf_path)}")
    else:
        print("PDF 파일을 찾을 수 없습니다.")
        return False

    # Gmail SMTP 서버 연결 및 전송
    try:
        # 앱 비밀번호 가져오기
        password = None

        # 1. 환경변수에서 확인
        password = os.environ.get('GMAIL_APP_PASSWORD')

        # 2. .env 파일에서 확인
        if not password:
            env_file = get_env_file_path()
            if env_file.exists():
                with open(env_file, 'r', encoding='utf-8') as f:
                    for line in f:
                        if line.startswith('GMAIL_APP_PASSWORD'):
                            password = line.split('=', 1)[1].strip().strip('\'"')
                            break

        # 3. 수동 입력
        if not password:
            print("\nGmail 앱 비밀번호가 필요합니다.")
            print("(2단계 인증이 활성화된 경우 앱 비밀번호를 사용하세요)")
            print("앱 비밀번호 생성: https://myaccount.google.com/apppasswords")
            password = getpass.getpass("Gmail 앱 비밀번호 입력: ")

        # Gmail SMTP 서버 연결
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, password)

        # 이메일 전송
        text = msg.as_string()
        server.sendmail(sender_email, recipients, text)
        server.quit()

        print(f"\n{GREEN}이메일 전송 완료!{RESET}")
        print(f"수신자: {', '.join(recipients)}")
        print(f"제목: {sunday_str} 행정광고")
        return True

    except Exception as e:
        print(f"\n이메일 전송 실패: {str(e)}")
        print("\n가능한 해결 방법:")
        print("1. 2단계 인증을 활성화하고 앱 비밀번호를 생성하세요")
        print("2. 환경변수 GMAIL_APP_PASSWORD를 설정하세요")
        print("3. Gmail 보안 설정을 확인하세요")
        return False

def main():
    """메인 실행 함수"""
    print("=" * 50)
    print("행정광고 PPT 내보내기 및 이메일 전송")
    print("=" * 50 + "\n")

    # 디렉토리 설정
    parent_dir = get_base_dir()
    admin_dir = os.path.join(parent_dir, "행정광고")

    if not os.path.exists(admin_dir):
        print(f"오류: {admin_dir} 디렉토리를 찾을 수 없습니다.")
        return False

    # 이번주 일요일 날짜
    this_sunday = get_this_sunday()
    sunday_str = this_sunday.strftime("%m%d")

    print(f"이번주 일요일: {this_sunday.strftime('%Y년 %m월 %d일')} ({sunday_str})")

    # PPT 파일 경로
    ppt_name = f"{sunday_str} 행정광고.pptx"
    ppt_path = os.path.join(admin_dir, ppt_name)

    if not os.path.exists(ppt_path):
        print(f"\n오류: {ppt_name} 파일을 찾을 수 없습니다.")
        print("먼저 weekly_announcement_ppt_copy.py를 실행하여 PPT를 생성하세요.")
        return False

    print(f"PPT 파일 발견: {ppt_name}")

    # PDF, PNG로 변환
    pdf_path, png_dir = export_ppt_to_pdf_and_png(ppt_path, admin_dir)

    if not pdf_path:
        print("\nPDF 변환에 실패했습니다.")
        return False

    # 수신자 선택
    recipients = select_recipients(RECIPIENTS)
    if not recipients:
        print(f"\n{YELLOW}수신자를 선택하지 않았습니다.{RESET}")
        return False

    # 이메일 전송 (PDF만)
    success = send_email_with_attachments(pdf_path, sunday_str, recipients)

    if success:
        print("\n모든 작업이 성공적으로 완료되었습니다!")
        print(f"- PDF: {os.path.basename(pdf_path)}")
        print(f"- PNG: {os.path.basename(png_dir)}/")
        print(f"- 이메일: {', '.join(recipients)}로 전송 완료")
    else:
        print("\n이메일 전송은 실패했지만 PDF, PNG는 생성되었습니다.")
        print(f"- PDF: {os.path.basename(pdf_path)}")
        print(f"- PNG: {os.path.basename(png_dir)}/")

    return True

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
