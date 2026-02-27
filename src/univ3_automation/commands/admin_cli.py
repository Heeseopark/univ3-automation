#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
행정팀 통합 CLI 메뉴
모든 행정 자동화 기능을 하나의 인터페이스에서 관리합니다.
"""

import os
import subprocess
import sys
from datetime import datetime, timedelta
from io import StringIO
from typing import TypedDict

import pytz

# 키보드 입력 처리
try:
    import msvcrt

    def get_key():
        """Windows에서 키 입력 받기"""
        key = msvcrt.getch()
        if key in [b"\x00", b"\xe0"]:  # 특수 키
            key = msvcrt.getch()
            if key == b"H":  # 위 화살표
                return "UP"
            elif key == b"P":  # 아래 화살표
                return "DOWN"
            elif key == b"K":  # 왼쪽 화살표
                return "LEFT"
            elif key == b"M":  # 오른쪽 화살표
                return "RIGHT"
        elif key == b"\r":  # Enter
            return "ENTER"
        elif key == b"\x1b":  # ESC
            return "ESC"
        return None

except ImportError:
    # Unix/Linux/Mac 환경
    import select
    import termios
    import tty

    def get_key():
        """Unix/Linux/Mac에서 키 입력 받기"""
        fd = sys.stdin.fileno()
        old_settings = termios.tcgetattr(fd)
        try:
            tty.setraw(fd)
            key = sys.stdin.read(1)
            if key == "\x1b":  # ESC sequence
                if select.select([sys.stdin], [], [], 0.01)[0]:
                    next_key = sys.stdin.read(1)
                    if next_key == "[" and select.select([sys.stdin], [], [], 0.01)[0]:
                        final_key = sys.stdin.read(1)
                        if final_key == "A":
                            return "UP"
                        elif final_key == "B":
                            return "DOWN"
                        elif final_key == "D":
                            return "LEFT"
                        elif final_key == "C":
                            return "RIGHT"
                return "ESC"
            elif key in ("\r", "\n"):
                return "ENTER"
        finally:
            termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)
        return None


class MenuItem(TypedDict):
    title: str
    script: str


class TabSection(TypedDict):
    label: str
    items: list[MenuItem]


def get_this_sunday():
    """현재 한국 시간 기준으로 이번 주 일요일 날짜를 반환"""
    kst = pytz.timezone("Asia/Seoul")
    now_kst = datetime.now(kst)
    weekday = now_kst.weekday()

    if weekday == 6:  # 이미 일요일인 경우
        sunday = now_kst
    else:
        sunday = now_kst + timedelta(days=6 - weekday)

    return sunday


def clear_screen():
    """화면 초기화"""
    os.system("cls" if os.name == "nt" else "clear")


def build_sections(mmdd: str) -> list[TabSection]:
    return [
        {
            "label": "집회준비",
            "items": [
                {"title": "프린트용 PDF 선택 + 이메일 전송", "script": "select_and_send_pdfs.py"},
                {"title": "큐티 월 업데이트 (엑셀 수정 + PDF)", "script": "update_qt_month.py"},
                {"title": "다음 주 출석부 생성 (구글 시트 → PDF)", "script": "attendance_generator.py"},
            ],
        },
        {
            "label": "카톡플친",
            "items": [
                {"title": f"{mmdd} 카톡플친 썸네일 생성", "script": "kakao_plus_friend_thumbnail_generator.py"},
                {"title": f"{mmdd} 카카오톡 플친 템플릿 생성", "script": "create_kakao_friend_template.py"},
                {"title": f"{mmdd} 카톡플친 PPT 열기", "script": "open_kakao_plus_friend.py"},
                {"title": "유튜브 재생목록 검색/링크 복사 (대학3부 주일집회 콘티)", "script": "youtube_playlist_search.py"},
            ],
        },
        {
            "label": "주보",
            "items": [
                {"title": f"{mmdd} 주보 열기 (자동 이동 포함)", "script": "open_weekly_bulletin.py"},
                {"title": f"{mmdd} 주보 PPT → PDF/PNG 변환", "script": "convert_weekly_bulletin.py"},
                {"title": f"{mmdd} 주보 카톡 채널 업로드", "script": "upload_bulletin_kakao.py"},
                {"title": f"{mmdd} 주보 PPT 카카오톡 전송 (GUI)", "script": "send_kakao_bulletin_ppt.py"},
                {"title": "주보 생일자 명단 추출 (행정파일)", "script": "birthday_selector.py"},
            ],
        },
        {
            "label": "행정광고",
            "items": [
                {"title": f"{mmdd} 행정광고 PPT 생성", "script": "create_admin_ppt.py"},
                {"title": f"{mmdd} 행정광고 열기", "script": "open_admin_announcement.py"},
                {"title": f"{mmdd} 연합광고 열기", "script": "open_union_announcement.py"},
                {"title": f"{mmdd} 연합광고 반영 (자동 이동 포함)", "script": "merge_union_announcement.py"},
                {"title": f"{mmdd} 행정광고 텍스트 박스 추가", "script": "add_text_box_to_ppt.py"},
                {"title": f"{mmdd} 행정광고 PPT 카카오톡 전송 (GUI)", "script": "send_kakao_admin_ppt.py"},
                {"title": f"{mmdd} 행정광고 PDF/PNG 추출 및 PDF 이메일 전송", "script": "export_and_send_announcement.py"},
                {"title": f"{mmdd} 행정광고 PNG 카카오톡 전송 (GUI)", "script": "send_kakao_announcement_images.py"},
                {"title": "행정광고 텍스트 선택 복사", "script": "text_line_selector.py"},
                {"title": "이미지에서 폰트 찾기", "script": "find_font_from_image.py"},
            ],
        },
        {
            "label": "기타",
            "items": [
                {"title": "다운로드 PDF 옮기기 (최근 5개 선택)", "script": "move_downloaded_pdfs.py"},
                {"title": "다운로드 이미지 옮기기 (최근 5개 선택)", "script": "move_downloaded_images.py"},
                {"title": "다운로드 PDF 카카오톡 전송 (최근 5개 선택)", "script": "send_kakao_downloaded_pdf.py"},
                {"title": "다운로드 PDF 옮기기 + PNG 변환", "script": "move_and_convert_pdfs.py"},
                {"title": "나가기", "script": "__exit__"},
            ],
        },
    ]


def _render_buffered_content(content: str) -> None:
    for line in content.split("\n"):
        sys.stdout.write(line + "\033[K\n")
    # 아래 영역을 지워 이전 프레임 잔상이 남지 않게 한다.
    sys.stdout.write("\033[J")
    sys.stdout.flush()


def display_tabbed_menu(
    sections: list[TabSection],
    tab_index: int,
    item_index: int,
    sunday_str: str,
    mmdd: str,
    first_draw: bool = False,
) -> None:
    """탭 기반 메뉴를 인플레이스 방식으로 렌더링한다."""
    if first_draw:
        clear_screen()
    else:
        # 전체 clear 대신 홈으로 이동해 같은 영역에 덮어쓴다.
        sys.stdout.write("\033[H")

    max_items = max(len(section["items"]) for section in sections)
    current_section = sections[tab_index]

    try:
        from rich.console import Console
        from rich.panel import Panel
        from rich.text import Text

        output = StringIO()
        console = Console(file=output, force_terminal=True)

        header = Text()
        header.append("행정팀 자동화 시스템\n", style="bold cyan")
        header.append(f"돌아오는 주일: {sunday_str}\n", style="cyan")
        header.append(f"기준 MMDD: {mmdd}", style="dim")
        console.print(Panel(header, border_style="cyan"))

        tab_line = "  ".join(
            f"[black on bright_cyan] {section['label']} [/]"
            if i == tab_index
            else f"[white] {section['label']} [/white]"
            for i, section in enumerate(sections)
        )
        console.print(tab_line)
        console.print("[cyan]" + "─" * 80 + "[/cyan]")

        for i in range(max_items):
            if i < len(current_section["items"]):
                item_title = current_section["items"][i]["title"]
                if i == item_index:
                    console.print(f"[bold cyan]▶ {item_title}[/bold cyan]")
                else:
                    console.print(f"  {item_title}")
            else:
                console.print("")

        console.print("[cyan]" + "─" * 80 + "[/cyan]")
        console.print("[yellow]←/→ 탭 이동 | ↑/↓ 항목 이동 | Enter: 선택 | ESC: 나가기[/yellow]")
        _render_buffered_content(output.getvalue())

    except ImportError:
        lines = [
            "=" * 80,
            " 행정팀 자동화 시스템",
            f" 돌아오는 주일: {sunday_str}",
            f" 기준 MMDD: {mmdd}",
            "=" * 80,
            "  ".join(
                f"[{section['label']}]" if i == tab_index else section["label"]
                for i, section in enumerate(sections)
            ),
            "─" * 80,
        ]

        for i in range(max_items):
            if i < len(current_section["items"]):
                item_title = current_section["items"][i]["title"]
                lines.append(f"{'▶ ' if i == item_index else '  '}{item_title}")
            else:
                lines.append("")

        lines.append("─" * 80)
        lines.append("←/→ 탭 이동 | ↑/↓ 항목 이동 | Enter: 선택 | ESC: 나가기")
        _render_buffered_content("\n".join(lines))


def run_script(script_name: str, script_dir: str):
    """스크립트 실행 - Python 파일 직접 실행"""
    try:
        from rich.console import Console
        from rich.panel import Panel

        console = Console()
        console.clear()

        console.print(Panel(f"[bold cyan]실행 중: {script_name}[/bold cyan]", border_style="cyan"))
        console.print()

    except ImportError:
        clear_screen()
        print(f"실행 중: {script_name}\n")
        print("=" * 60)

    script_path = os.path.join(script_dir, script_name)

    venv_python = os.path.join(os.path.dirname(script_dir), ".venv", "Scripts", "python.exe")
    if not os.path.exists(venv_python):
        venv_python = sys.executable

    try:
        result = subprocess.run(
            [venv_python, script_path],
            cwd=script_dir,
            capture_output=False,
            text=True,
        )

        if result.returncode == 0:
            try:
                from rich.console import Console

                console = Console()
                console.print("\n[bold green]성공적으로 완료되었습니다![/bold green]")
            except ImportError:
                print("\n성공적으로 완료되었습니다!")
        else:
            print(f"\n오류가 발생했습니다. (코드: {result.returncode})")

    except FileNotFoundError:
        print(f"\n오류: {script_name} 파일을 찾을 수 없습니다.")
    except Exception as e:
        print(f"\n오류 발생: {str(e)}")

    print("\n아무 키나 누르면 메뉴로 돌아갑니다...")
    if os.name == "nt":
        msvcrt.getch()
    else:
        input()


def print_exit_message():
    try:
        from rich.console import Console

        console = Console()
        console.clear()
        console.print("[bold cyan]행정팀 자동화 시스템을 종료합니다.[/bold cyan]")
        console.print("좋은 하루 되세요!")
    except ImportError:
        clear_screen()
        print("행정팀 자동화 시스템을 종료합니다.")
        print("좋은 하루 되세요!")


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    selected_tab_index = 0
    selected_item_index = 0
    first_draw = True

    while True:
        sunday = get_this_sunday()
        sunday_str = sunday.strftime("%Y년 %m월 %d일 (%m%d)")
        mmdd = sunday.strftime("%m%d")
        sections = build_sections(mmdd)

        if selected_tab_index >= len(sections):
            selected_tab_index = 0
        current_items = sections[selected_tab_index]["items"]
        if selected_item_index >= len(current_items):
            selected_item_index = 0

        display_tabbed_menu(
            sections=sections,
            tab_index=selected_tab_index,
            item_index=selected_item_index,
            sunday_str=sunday_str,
            mmdd=mmdd,
            first_draw=first_draw,
        )
        first_draw = False

        key = get_key()
        current_items = sections[selected_tab_index]["items"]

        if key == "LEFT":
            if selected_tab_index > 0:
                selected_tab_index -= 1
                selected_item_index = 0
        elif key == "RIGHT":
            if selected_tab_index < len(sections) - 1:
                selected_tab_index += 1
                selected_item_index = 0
        elif key == "UP":
            if selected_item_index > 0:
                selected_item_index -= 1
        elif key == "DOWN":
            if selected_item_index < len(current_items) - 1:
                selected_item_index += 1
        elif key == "ENTER":
            selected_item = current_items[selected_item_index]
            script_name = selected_item["script"]
            if script_name == "__exit__":
                print_exit_message()
                sys.exit(0)

            run_script(script_name, script_dir)
            first_draw = True
        elif key == "ESC":
            print_exit_message()
            sys.exit(0)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        try:
            from rich.console import Console

            console = Console()
            console.clear()
            console.print("\n[bold cyan]프로그램을 종료합니다.[/bold cyan]")
        except ImportError:
            clear_screen()
            print("\n프로그램을 종료합니다.")
        sys.exit(0)
