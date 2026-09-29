#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF 선택 및 이메일 전송 스크립트
주보 및 프린트 폴더의 PDF 파일을 선택하여 이메일로 전송합니다.
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
import smtplib
from datetime import datetime, timedelta
import pytz
from pathlib import Path
import getpass
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText

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

# 이메일 수신자 설정
RECIPIENTS = [
    "heeseopark99@gmail.com",
    "kangyeso0204@gmail.com",
    "blessolivia060606@gmail.com"
    # 필요시 추가 수신자를 여기에 추가
]

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

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def find_pdf_files():
    """주보, 프린트, 군지체, 리더파일 폴더에서 PDF 파일을 찾아 반환"""
    base_dir = Path(get_base_dir())

    pdf_files = []
    warnings = []

    # 1. 주보 폴더에서 이번 주 일요일 PDF 찾기
    jubo_dir = base_dir / "주보"
    if jubo_dir.exists():
        sunday = get_this_sunday()
        sunday_str = sunday.strftime("%m%d")

        jubo_found = False
        for pdf_file in jubo_dir.glob("*.pdf"):
            if sunday_str in pdf_file.name:
                pdf_files.append({
                    'path': str(pdf_file),
                    'name': pdf_file.name,
                    'folder': '주보',
                    'type': 'jubo',
                    'size': pdf_file.stat().st_size
                })
                jubo_found = True

        if not jubo_found:
            warnings.append(f"{sunday_str} 주보 PDF가 없습니다.")

    # 2. 프린트 폴더에서 모든 PDF 찾기
    print_dir = base_dir / "프린트"
    if print_dir.exists():
        for pdf_file in print_dir.glob("*.pdf"):
            pdf_files.append({
                'path': str(pdf_file),
                'name': pdf_file.name,
                'folder': '프린트',
                'type': 'print',
                'size': pdf_file.stat().st_size
            })

    # 3. 프린트/교재 폴더에서 PDF 찾기
    print_gyojae_dir = print_dir / "교재" if print_dir.exists() else None
    if print_gyojae_dir and print_gyojae_dir.exists():
        for pdf_file in print_gyojae_dir.glob("*.pdf"):
            pdf_files.append({
                'path': str(pdf_file),
                'name': pdf_file.name,
                'folder': '프린트',
                'type': 'gyojae',
                'size': pdf_file.stat().st_size
            })

    # 4. 군지체 폴더에서 PDF 찾기
    gunjiche_dir = base_dir / "군지체"
    if gunjiche_dir.exists():
        for pdf_file in gunjiche_dir.glob("*.pdf"):
            pdf_files.append({
                'path': str(pdf_file),
                'name': pdf_file.name,
                'folder': '군지체',
                'type': 'gunjiche',
                'size': pdf_file.stat().st_size
            })

    # 5. 리더파일 폴더에서 PDF 찾기
    reader_dir = base_dir / "리더파일"
    if reader_dir.exists():
        for pdf_file in reader_dir.glob("*.pdf"):
            pdf_files.append({
                'path': str(pdf_file),
                'name': pdf_file.name,
                'folder': '리더파일',
                'type': 'reader',
                'size': pdf_file.stat().st_size
            })

    return pdf_files, warnings

def format_size(size):
    """파일 크기를 읽기 쉬운 형식으로 변환"""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size < 1024.0:
            return f"{size:.1f} {unit}"
        size /= 1024.0
    return f"{size:.1f} TB"

def display_menu(pdf_files, selected, current_index, warnings=None):
    clear_screen()
    print("=" * 70)
    print(f"{CYAN}PDF 파일 선택 및 이메일 전송{RESET}")
    print("=" * 70)
    print("↑/↓: 이동, 스페이스: 선택/해제, Enter: 전송, ESC: 취소")
    print("=" * 70)
    print()

    # 모든 폴더를 순서대로 표시
    folders = ['주보', '프린트', '군지체', '리더파일']

    for folder in folders:
        print(f"\n{GREEN}[{folder}]{RESET}")
        print("-" * 50)

        # 해당 폴더의 PDF 파일들 찾기
        folder_pdfs = [(i, pdf) for i, pdf in enumerate(pdf_files) if pdf['folder'] == folder]

        if folder_pdfs:
            for i, pdf in folder_pdfs:
                # 선택 표시
                selected_mark = f"{MAGENTA}[✓]{RESET} " if i in selected else "[ ] "

                # 커서 표시
                cursor = "▶ " if i == current_index else "  "

                # 파일 정보 표시
                file_info = f"{pdf['name']} ({format_size(pdf['size'])})"

                if i in selected:
                    print(f"{cursor}{selected_mark}{MAGENTA}{file_info}{RESET}")
                elif i == current_index:
                    print(f"{cursor}{selected_mark}{CYAN}{file_info}{RESET}")
                else:
                    print(f"{cursor}{selected_mark}{file_info}")
        else:
            # 비어있는 폴더 처리
            if folder == '주보' and warnings:
                # 주보 폴더에 경고 메시지 표시
                for warning in warnings:
                    print(f"  {YELLOW}⚠️  {warning}{RESET}")
            else:
                print(f"  {YELLOW}(비어있음){RESET}")

    print()
    print("=" * 70)
    print(f"선택된 파일: {len(selected)}개")

    if selected:
        total_size = sum(pdf_files[i]['size'] for i in selected)
        print(f"총 크기: {format_size(total_size)}")

def input_print_info(pdf_files, selected_indices):
    """선택된 PDF 파일별로 프린트 정보 입력받기"""
    print_info = {}

    clear_screen()
    print("=" * 70)
    print(f"{CYAN}프린트 정보 입력{RESET}")
    print("=" * 70)
    print("각 파일별로 프린트 정보를 입력하세요 (예: 180장, 1부)")
    print("입력하지 않으려면 Enter만 누르세요")
    print("=" * 70)
    print()

    for idx in selected_indices:
        pdf = pdf_files[idx]
        print(f"\n{GREEN}[{pdf['folder']}]{RESET} {pdf['name']}")
        print("-" * 50)
        info = input("프린트 정보 입력: ").strip()
        print_info[idx] = info

    return print_info

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

def send_email_with_pdfs(pdf_files, selected_indices, recipients, print_info):
    """선택된 PDF를 이메일로 전송"""
    print(f"\n{CYAN}이메일 전송 준비...{RESET}")

    sunday = get_this_sunday()
    date_str = sunday.strftime("%m%d")

    # 발신자 이메일 (첫 번째 수신자와 동일)
    sender_email = recipients[0]

    # 이메일 구성
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ', '.join(recipients)
    msg['Subject'] = f"{date_str} 프린트"

    # HTML 본문 작성
    body_html = """
    <html>
      <body>
        <p><strong>프린트 장수</strong></p>
        <ul>
    """

    # PDF 첨부
    attached_count = 0
    for idx in selected_indices:
        pdf = pdf_files[idx]
        pdf_path = pdf['path']

        if os.path.exists(pdf_path):
            try:
                with open(pdf_path, "rb") as f:
                    part = MIMEApplication(f.read(), _subtype='pdf')
                    part.add_header(
                        'Content-Disposition',
                        'attachment',
                        filename=pdf['name']
                    )
                    msg.attach(part)

                    # 파일명에서 .pdf 제거
                    file_display_name = pdf['name'].replace('.pdf', '').replace('.PDF', '')

                    # 프린트 정보가 있으면 추가
                    if idx in print_info and print_info[idx]:
                        body_html += f"          <li>{file_display_name} - {print_info[idx]}</li>\n"
                    else:
                        body_html += f"          <li>{file_display_name}</li>\n"

                    attached_count += 1
                    print(f"  - {pdf['name']} 첨부 완료")
            except Exception as e:
                print(f"  - {pdf['name']} 첨부 실패: {str(e)}")

    body_html += """
        </ul>
      </body>
    </html>
    """

    msg.attach(MIMEText(body_html, 'html', 'utf-8'))

    if attached_count == 0:
        print(f"{YELLOW}첨부할 파일이 없습니다.{RESET}")
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
            print(f"\n{YELLOW}Gmail 앱 비밀번호가 필요합니다.{RESET}")
            print("(2단계 인증이 활성화된 경우 앱 비밀번호를 사용하세요)")
            print("앱 비밀번호 생성: https://myaccount.google.com/apppasswords")
            password = getpass.getpass("Gmail 앱 비밀번호 입력: ")

        # Gmail SMTP 서버 연결
        print(f"\n{CYAN}이메일 전송 중...{RESET}")
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, password)

        # 이메일 전송
        text = msg.as_string()
        server.sendmail(sender_email, recipients, text)
        server.quit()

        print(f"\n{GREEN}이메일 전송 완료!{RESET}")
        print(f"수신자: {', '.join(recipients)}")
        print(f"제목: {msg['Subject']}")
        print(f"첨부 파일: {attached_count}개")
        return True

    except Exception as e:
        print(f"\n{YELLOW}이메일 전송 실패: {str(e)}{RESET}")
        print("\n가능한 해결 방법:")
        print("1. 2단계 인증을 활성화하고 앱 비밀번호를 생성하세요")
        print("2. 환경변수 GMAIL_APP_PASSWORD를 설정하세요")
        print("3. Gmail 보안 설정을 확인하세요")
        return False

def main():
    """메인 실행 함수"""

    # PDF 파일 찾기
    pdf_files, warnings = find_pdf_files()

    if not pdf_files:
        print(f"{YELLOW}PDF 파일을 찾을 수 없습니다.{RESET}")
        print("\n주보 폴더나 프린트 폴더에 PDF 파일이 있는지 확인하세요.")
        return

    # PDF 선택 UI
    selected = set()
    current_index = 0

    try:
        if os.name == 'nt':
            import msvcrt

            while True:
                display_menu(pdf_files, selected, current_index, warnings)

                key = msvcrt.getch()

                if key == b'\xe0' or key == b'\x00':
                    key = msvcrt.getch()
                    if key == b'H':  # 위쪽 화살표
                        current_index = max(0, current_index - 1)
                    elif key == b'P':  # 아래쪽 화살표
                        current_index = min(len(pdf_files) - 1, current_index + 1)

                elif key == b' ':  # 스페이스바
                    if current_index in selected:
                        selected.remove(current_index)
                    else:
                        selected.add(current_index)

                elif key == b'\r':  # Enter
                    if selected:
                        break
                    else:
                        print(f"\n{YELLOW}파일을 선택해주세요.{RESET}")
                        msvcrt.getch()

                elif key == b'\x1b':  # ESC
                    print(f"\n{YELLOW}작업이 취소되었습니다.{RESET}")
                    return

        else:
            # Unix/Linux/Mac 환경
            import termios, tty

            old_settings = termios.tcgetattr(sys.stdin)
            try:
                tty.setraw(sys.stdin.fileno())

                while True:
                    display_menu(pdf_files, selected, current_index, warnings)

                    key = sys.stdin.read(1)

                    if key == '\x1b':
                        next_key = sys.stdin.read(2)
                        if next_key == '[A':  # 위쪽 화살표
                            current_index = max(0, current_index - 1)
                        elif next_key == '[B':  # 아래쪽 화살표
                            current_index = min(len(pdf_files) - 1, current_index + 1)
                        elif next_key == '':
                            print(f"\n{YELLOW}작업이 취소되었습니다.{RESET}")
                            return

                    elif key == ' ':  # 스페이스바
                        if current_index in selected:
                            selected.remove(current_index)
                        else:
                            selected.add(current_index)

                    elif key == '\r':  # Enter
                        if selected:
                            break

            finally:
                termios.tcsetattr(sys.stdin, termios.TCSADRAIN, old_settings)

    except KeyboardInterrupt:
        print(f"\n\n{YELLOW}작업이 중단되었습니다.{RESET}")
        return

    # 프린트 정보 입력받기
    print_info = input_print_info(pdf_files, sorted(selected))

    # 수신자 선택
    recipients = select_recipients(RECIPIENTS)
    if not recipients:
        print(f"\n{YELLOW}수신자를 선택하지 않았습니다.{RESET}")
        return

    # 이메일 전송
    if selected:
        clear_screen()
        print("=" * 70)
        print(f"{CYAN}선택된 파일:{RESET}")
        print("=" * 70)

        for idx in sorted(selected):
            pdf = pdf_files[idx]
            if idx in print_info and print_info[idx]:
                print(f"  - {pdf['name']} ({pdf['folder']}) - {print_info[idx]}")
            else:
                print(f"  - {pdf['name']} ({pdf['folder']})")

        print(f"\n수신자: {', '.join(recipients)}")

        if send_email_with_pdfs(pdf_files, sorted(selected), recipients, print_info):
            print(f"\n{GREEN}모든 작업이 완료되었습니다!{RESET}")
        else:
            print(f"\n{YELLOW}이메일 전송에 실패했습니다.{RESET}")

if __name__ == "__main__":
    main()
