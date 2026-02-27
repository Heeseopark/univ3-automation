#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
유튜브 재생목록 검색 및 링크 복사
'대학3부 주일집회 콘티' 재생목록을 검색하여 선택하고, 1번 인덱스 링크를 복사합니다.
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
from google.oauth2 import service_account
from googleapiclient.discovery import build
import pyperclip

# Google API 인증 설정
SCOPES = ['https://www.googleapis.com/auth/youtube.readonly']

# 색상 코드 정의
try:
    import colorama
    colorama.init()
    MAGENTA = '\033[95m'
    CYAN = '\033[96m'
    WHITE = '\033[97m'
    YELLOW = '\033[93m'
    GREEN = '\033[92m'
    RED = '\033[91m'
    RESET = '\033[0m'
except ImportError:
    MAGENTA = CYAN = WHITE = YELLOW = GREEN = RED = RESET = ''

def clear_screen():
    """화면 초기화"""
    os.system('cls' if os.name == 'nt' else 'clear')

def get_youtube_service():
    """YouTube API 서비스 객체 생성"""
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

    # YouTube API 서비스 생성
    youtube = build('youtube', 'v3', credentials=creds)
    return youtube

def search_playlists(youtube, query, channel_filter=None):
    """재생목록 검색 - 모든 결과 반환"""
    try:
        playlists = []
        next_page_token = None

        # 모든 페이지 검색 (최대 50개씩, 여러 페이지)
        while True:
            request = youtube.search().list(
                part="snippet",
                q=query,
                type="playlist",
                maxResults=50,
                pageToken=next_page_token
            )
            response = request.execute()

            for item in response.get('items', []):
                playlist_id = item['id']['playlistId']
                title = item['snippet']['title']
                description = item['snippet']['description']
                channel_title = item['snippet']['channelTitle']

                # 채널 필터링
                if channel_filter and channel_filter not in channel_title:
                    continue

                playlists.append({
                    'id': playlist_id,
                    'title': title,
                    'description': description,
                    'channel': channel_title
                })

            # 다음 페이지가 있으면 계속, 없으면 종료
            next_page_token = response.get('nextPageToken')
            if not next_page_token:
                break

        return playlists

    except Exception as e:
        print(f"{RED}✗ 재생목록 검색 오류: {e}{RESET}")
        return []

def get_playlist_items(youtube, playlist_id):
    """재생목록의 동영상 목록 가져오기"""
    try:
        request = youtube.playlistItems().list(
            part="snippet,contentDetails",
            playlistId=playlist_id,
            maxResults=50
        )
        response = request.execute()

        items = []
        for item in response.get('items', []):
            video_id = item['contentDetails']['videoId']
            title = item['snippet']['title']
            items.append({
                'video_id': video_id,
                'title': title
            })

        return items

    except Exception as e:
        print(f"{RED}✗ 재생목록 항목 조회 오류: {e}{RESET}")
        return []

def display_playlist_menu(playlists, selected_index, scroll_offset=0):
    """재생목록 선택 메뉴 표시 (스크롤 가능)"""
    clear_screen()
    print("=" * 80)
    print(f"{CYAN}유튜브 재생목록 검색 결과 (총 {len(playlists)}개){RESET}")
    print("=" * 80)
    print("↑/↓: 이동, Enter: 선택, ESC: 취소")
    print("=" * 80)
    print()

    # 화면에 표시할 최대 항목 수
    max_display = 15
    start_idx = scroll_offset
    end_idx = min(scroll_offset + max_display, len(playlists))

    for i in range(start_idx, end_idx):
        playlist = playlists[i]
        cursor = "▶ " if i == selected_index else "  "

        if i == selected_index:
            print(f"{cursor}{CYAN}{i+1}. {playlist['title']}{RESET}")
            print(f"     채널: {playlist['channel']}")
        else:
            print(f"{cursor}{i+1}. {playlist['title']}")

    print()
    print("=" * 80)
    if len(playlists) > max_display:
        print(f"{YELLOW}표시: {start_idx + 1}-{end_idx} / {len(playlists)}{RESET}")

def main():
    """메인 함수"""
    try:
        print("=" * 80)
        print("유튜브 재생목록 검색 및 링크 복사")
        print("=" * 80)
        print()

        # YouTube API 서비스 생성
        print("YouTube API 연결 중...")
        youtube = get_youtube_service()
        print(f"{GREEN}✓ YouTube API 연결 완료{RESET}\n")

        # 재생목록 검색
        query = "대학3부 주일집회 콘티"
        channel_filter = "이주은"  # 이주은 채널로 필터링
        print(f"'{query}' 검색 중 (채널 필터: {channel_filter})...")
        playlists = search_playlists(youtube, query, channel_filter)

        if not playlists:
            print(f"{RED}✗ 검색 결과가 없습니다.{RESET}")
            return 1

        print(f"{GREEN}✓ {len(playlists)}개의 재생목록을 찾았습니다.{RESET}\n")

        # 재생목록 선택 UI
        selected_index = 0
        scroll_offset = 0
        max_display = 15

        if os.name == 'nt':
            import msvcrt

            while True:
                display_playlist_menu(playlists, selected_index, scroll_offset)

                key = msvcrt.getch()

                if key == b'\xe0' or key == b'\x00':
                    key = msvcrt.getch()
                    if key == b'H':  # 위쪽 화살표
                        selected_index = max(0, selected_index - 1)
                        # 스크롤 조정
                        if selected_index < scroll_offset:
                            scroll_offset = selected_index
                    elif key == b'P':  # 아래쪽 화살표
                        selected_index = min(len(playlists) - 1, selected_index + 1)
                        # 스크롤 조정
                        if selected_index >= scroll_offset + max_display:
                            scroll_offset = selected_index - max_display + 1

                elif key == b'\r':  # Enter
                    break

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
                    display_playlist_menu(playlists, selected_index, scroll_offset)

                    key = sys.stdin.read(1)

                    if key == '\x1b':
                        next_key = sys.stdin.read(2)
                        if next_key == '[A':  # 위쪽 화살표
                            selected_index = max(0, selected_index - 1)
                            # 스크롤 조정
                            if selected_index < scroll_offset:
                                scroll_offset = selected_index
                        elif next_key == '[B':  # 아래쪽 화살표
                            selected_index = min(len(playlists) - 1, selected_index + 1)
                            # 스크롤 조정
                            if selected_index >= scroll_offset + max_display:
                                scroll_offset = selected_index - max_display + 1
                        elif next_key == '':
                            print("\n작업이 취소되었습니다.")
                            return 0

                    elif key == '\r':  # Enter
                        break

            finally:
                termios.tcsetattr(sys.stdin, termios.TCSADRAIN, old_settings)

        # 선택된 재생목록
        selected_playlist = playlists[selected_index]
        playlist_id = selected_playlist['id']

        clear_screen()
        print("=" * 80)
        print(f"{CYAN}선택된 재생목록: {selected_playlist['title']}{RESET}")
        print("=" * 80)
        print()

        # 재생목록 항목 가져오기
        print("재생목록 항목 조회 중...")
        items = get_playlist_items(youtube, playlist_id)

        if not items:
            print(f"{RED}✗ 재생목록이 비어있습니다.{RESET}")
            return 1

        if len(items) < 2:
            print(f"{RED}✗ 재생목록에 동영상이 1개 이하입니다. (최소 2개 필요){RESET}")
            return 1

        print(f"{GREEN}✓ {len(items)}개의 동영상을 찾았습니다.{RESET}\n")

        # 1번 인덱스(두 번째 동영상) 링크 생성
        video_id = items[1]['video_id']
        video_title = items[1]['title']
        video_link = f"https://www.youtube.com/watch?v={video_id}&list={playlist_id}&index=2"

        print("=" * 80)
        print(f"{YELLOW}1번 인덱스 동영상:{RESET}")
        print(f"  제목: {video_title}")
        print(f"  링크: {video_link}")
        print("=" * 80)
        print()

        # 클립보드에 복사
        try:
            pyperclip.copy(video_link)
            print(f"{GREEN}✓ 링크가 클립보드에 복사되었습니다!{RESET}")
        except Exception as e:
            print(f"{YELLOW}⚠ 클립보드 복사 실패: {e}{RESET}")
            print("링크를 수동으로 복사해주세요.")

        print()
        print("=" * 80)
        print(f"{GREEN}✓ 작업이 완료되었습니다!{RESET}")
        print("=" * 80)

        return 0

    except Exception as e:
        print(f"\n{RED}✗ 오류 발생: {e}{RESET}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())

