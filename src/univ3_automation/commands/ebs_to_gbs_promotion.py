#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EBS -> GBS 등반 처리 스크립트.

이름으로 EBS 등반 대상자를 검색/선택한 뒤 라인업정보 소속과 No를 갱신하고,
출석부의 주간 출석 값/시각 서식을 새 No 위치로 이동한다.
"""

from __future__ import annotations

try:
    from runtime_paths import get_google_credentials_candidates, get_google_sheet_id
except ImportError:
    from univ3_automation.commands.runtime_paths import (
        get_google_credentials_candidates,
        get_google_sheet_id,
    )

import copy
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

LINEUP_SHEET = "라인업정보"
ATTENDANCE_SHEET = "출석부"
SPREADSHEET_ID = get_google_sheet_id()
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

EBS_GROUP_NUMBERS = (25, 45, 55, 75, 56, 10, 20, 30, 40, 50, 60, 70)
EBS_GROUP_NUMBER_SET = set(EBS_GROUP_NUMBERS)
IMMOVABLE_LEADER_ROLES = {"고을지기", "리더", "간사"}
ATTENDANCE_START_COLUMN_INDEX = 9   # J
ATTENDANCE_END_COLUMN_INDEX = 36    # AJ + 1
ATTENDANCE_CELL_WIDTH = ATTENDANCE_END_COLUMN_INDEX - ATTENDANCE_START_COLUMN_INDEX


class PromotionError(RuntimeError):
    """사용자에게 그대로 보여 줄 수 있는 검증/실행 오류."""


@dataclass(frozen=True)
class LineupMember:
    sheet_row: int
    no: int
    group_no: int
    town: str
    leader: str
    gender: str
    grade: str
    name: str
    leader_role: str

    @property
    def display_name(self) -> str:
        prefix = f"{self.gender}{self.grade}".strip()
        if prefix:
            return f"{prefix} {self.name}".strip()
        return self.name

    @property
    def is_movable(self) -> bool:
        return self.leader_role == ""


@dataclass(frozen=True)
class GroupOption:
    group_no: int
    town: str
    leader_name: str
    leader_role: str
    member_count: int

    @property
    def is_ebs(self) -> bool:
        return self.group_no in EBS_GROUP_NUMBER_SET


@dataclass(frozen=True)
class AttendanceRowRef:
    sheet_row: int
    order_no: int | None
    group_no: int | None
    key: int


@dataclass(frozen=True)
class PromotionPlan:
    source_member: LineupMember
    source_ebs: GroupOption
    target_group: GroupOption
    old_no: int
    new_no: int
    old_attendance_row: AttendanceRowRef
    new_attendance_row: AttendanceRowRef


@dataclass(frozen=True)
class AttendanceMergePlan:
    cleared_source_cells: list[dict[str, Any]]
    merged_destination_cells: list[dict[str, Any]]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def parse_int(value: Any) -> int | None:
    if value is None:
        return None

    text = str(value).strip()
    if not text:
        return None

    try:
        return int(float(text))
    except ValueError:
        return None


def cell_value(row: list[Any], index: int) -> Any:
    if index >= len(row):
        return None
    return row[index]


def find_credentials_file() -> Path:
    for candidate in get_google_credentials_candidates():
        if candidate.exists():
            return candidate
    raise PromotionError("Google 서비스 계정 인증 파일을 찾을 수 없습니다.")


def build_service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    credentials = service_account.Credentials.from_service_account_file(
        str(find_credentials_file()),
        scopes=SCOPES,
    )
    return build("sheets", "v4", credentials=credentials, cache_discovery=False)


def fetch_sheet_metadata(service) -> dict[str, int]:
    response = (
        service.spreadsheets()
        .get(spreadsheetId=SPREADSHEET_ID, fields="sheets(properties(sheetId,title))")
        .execute()
    )
    metadata: dict[str, int] = {}
    for sheet in response.get("sheets", []):
        properties = sheet.get("properties", {})
        title = properties.get("title")
        sheet_id = properties.get("sheetId")
        if title is not None and sheet_id is not None:
            metadata[str(title)] = int(sheet_id)
    return metadata


def fetch_values_snapshot(service) -> tuple[list[list[Any]], list[list[Any]]]:
    response = (
        service.spreadsheets()
        .values()
        .batchGet(
            spreadsheetId=SPREADSHEET_ID,
            ranges=[f"{LINEUP_SHEET}!A1:J", f"{ATTENDANCE_SHEET}!B6:AJ"],
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
    )
    ranges = response.get("valueRanges", [])
    if len(ranges) != 2:
        raise PromotionError("필수 시트 값을 읽지 못했습니다.")

    lineup_rows = ranges[0].get("values", [])
    attendance_rows = ranges[1].get("values", [])
    if not lineup_rows:
        raise PromotionError("라인업정보 시트가 비어 있습니다.")
    if not attendance_rows:
        raise PromotionError("출석부 시트가 비어 있습니다.")
    return lineup_rows, attendance_rows


def parse_lineup_members(rows: list[list[Any]]) -> list[LineupMember]:
    members: list[LineupMember] = []
    for sheet_row, row in enumerate(rows[1:], start=2):
        no = parse_int(cell_value(row, 0))
        group_no = parse_int(cell_value(row, 1))
        name = normalize_text(cell_value(row, 6))
        if no is None or group_no is None or not name:
            continue

        members.append(
            LineupMember(
                sheet_row=sheet_row,
                no=no,
                group_no=group_no,
                town=normalize_text(cell_value(row, 2)),
                leader=normalize_text(cell_value(row, 3)),
                gender=normalize_text(cell_value(row, 4)),
                grade=normalize_text(cell_value(row, 5)),
                name=name,
                leader_role=normalize_text(cell_value(row, 9)),
            )
        )

    return members


def parse_attendance_rows(rows: list[list[Any]]) -> dict[int, AttendanceRowRef]:
    refs: dict[int, AttendanceRowRef] = {}
    for sheet_row, row in enumerate(rows[1:], start=7):
        key = parse_int(cell_value(row, 2))
        if key is None:
            continue

        refs[key] = AttendanceRowRef(
            sheet_row=sheet_row,
            order_no=parse_int(cell_value(row, 0)),
            group_no=parse_int(cell_value(row, 1)),
            key=key,
        )

    return refs


def group_members_by_group_no(members: list[LineupMember]) -> dict[int, list[LineupMember]]:
    grouped: dict[int, list[LineupMember]] = {}
    for member in members:
        grouped.setdefault(member.group_no, []).append(member)
    for rows in grouped.values():
        rows.sort(key=lambda item: item.no)
    return grouped


def group_root_member(rows: list[LineupMember]) -> LineupMember:
    if not rows:
        raise PromotionError("비어 있는 그룹입니다.")

    expected_root_no = rows[0].group_no * 100 + 1
    for member in rows:
        if member.no == expected_root_no:
            return member
    return min(rows, key=lambda item: item.no)


def build_ebs_options(members: list[LineupMember]) -> list[GroupOption]:
    options: list[GroupOption] = []
    grouped_members = group_members_by_group_no(members)
    for group_no in EBS_GROUP_NUMBERS:
        rows = grouped_members.get(group_no)
        if not rows:
            continue

        root = group_root_member(rows)
        movable_count = sum(1 for row in rows if row.is_movable)
        if movable_count == 0:
            continue

        options.append(
            GroupOption(
                group_no=group_no,
                town=root.town,
                leader_name=root.name,
                leader_role=root.leader_role,
                member_count=movable_count,
            )
        )
    return options


def build_gbs_options(members: list[LineupMember], town: str) -> list[GroupOption]:
    options: list[GroupOption] = []
    for group_no, rows in sorted(group_members_by_group_no(members).items()):
        root = group_root_member(rows)
        if group_no <= 0:
            continue
        if root.town != town:
            continue
        if group_no in EBS_GROUP_NUMBER_SET:
            continue
        if root.leader_role == "고을지기":
            continue

        options.append(
            GroupOption(
                group_no=group_no,
                town=root.town,
                leader_name=root.name,
                leader_role=root.leader_role,
                member_count=len(rows),
            )
        )
    return options


def build_town_options(members: list[LineupMember]) -> list[str]:
    grouped = group_members_by_group_no(members)
    towns: list[str] = []
    seen: set[str] = set()
    for group_no, rows in sorted(grouped.items()):
        root = group_root_member(rows)
        if group_no <= 0:
            continue
        if group_no in EBS_GROUP_NUMBER_SET or root.leader_role == "고을지기":
            continue
        if root.town in seen:
            continue
        seen.add(root.town)
        towns.append(root.town)
    return towns


def build_member_options(members: list[LineupMember], ebs_group_no: int) -> list[LineupMember]:
    options = [
        member
        for member in members
        if member.group_no == ebs_group_no and member.is_movable
    ]
    options.sort(key=lambda item: item.no)
    return options


def build_searchable_ebs_members(members: list[LineupMember]) -> list[LineupMember]:
    options = [
        member
        for member in members
        if member.group_no in EBS_GROUP_NUMBER_SET and member.is_movable
    ]
    options.sort(key=lambda item: (item.name, item.group_no, item.no))
    return options


def normalize_search_query(value: str) -> str:
    return "".join(value.split()).lower()


def search_ebs_members_by_name(members: list[LineupMember], query: str) -> list[LineupMember]:
    query_key = normalize_search_query(query)
    if not query_key:
        return []

    matched: list[tuple[int, str, int, int, LineupMember]] = []
    for member in members:
        name_key = normalize_search_query(member.name)
        display_key = normalize_search_query(member.display_name)
        if query_key not in name_key and query_key not in display_key:
            continue

        if name_key == query_key:
            priority = 0
        elif name_key.startswith(query_key):
            priority = 1
        elif display_key.startswith(query_key):
            priority = 2
        else:
            priority = 3

        matched.append((priority, member.name, member.group_no, member.no, member))

    matched.sort(key=lambda item: (item[0], item[1], item[2], item[3]))
    return [item[4] for item in matched]


def find_ebs_option(options: list[GroupOption], group_no: int) -> GroupOption:
    for option in options:
        if option.group_no == group_no:
            return option
    raise PromotionError("선택한 대상자의 EBS 정보를 찾지 못했습니다.")


def build_promotion_plan(
    members: list[LineupMember],
    attendance_rows: dict[int, AttendanceRowRef],
    source_ebs: GroupOption,
    source_member: LineupMember,
    target_group: GroupOption,
) -> PromotionPlan:
    if source_ebs.group_no not in EBS_GROUP_NUMBER_SET:
        raise PromotionError("선택한 EBS 조번호가 허용된 EBS 번호가 아닙니다.")

    if source_member.group_no != source_ebs.group_no:
        raise PromotionError("선택한 조원이 선택한 EBS에 속해 있지 않습니다.")

    if source_member.leader_role in IMMOVABLE_LEADER_ROLES:
        raise PromotionError("리더 역할이 있는 인원은 등반 처리 대상에서 제외됩니다.")

    if target_group.group_no in EBS_GROUP_NUMBER_SET:
        raise PromotionError("대상 GBS가 아니라 EBS가 선택되었습니다.")

    if source_member.group_no == target_group.group_no:
        raise PromotionError("같은 GBS로는 등반 처리할 수 없습니다.")

    target_members = [member for member in members if member.group_no == target_group.group_no]
    if not target_members:
        raise PromotionError("대상 GBS에 기존 인원이 없어 새 No를 계산할 수 없습니다.")

    new_no = max(member.no for member in target_members) + 1
    existing_nos = {member.no for member in members}
    if new_no in existing_nos:
        raise PromotionError(f"새 No {new_no}가 기존 인원과 중복됩니다.")

    old_attendance_row = attendance_rows.get(source_member.no)
    if old_attendance_row is None:
        raise PromotionError("출석부에서 기존 No 위치를 찾지 못했습니다.")

    new_attendance_row = attendance_rows.get(new_no)
    if new_attendance_row is None:
        raise PromotionError("출석부에서 새 No 위치를 찾지 못했습니다.")

    return PromotionPlan(
        source_member=source_member,
        source_ebs=source_ebs,
        target_group=target_group,
        old_no=source_member.no,
        new_no=new_no,
        old_attendance_row=old_attendance_row,
        new_attendance_row=new_attendance_row,
    )


def fetch_attendance_row_cells(service, row_number: int) -> list[dict[str, Any]]:
    response = (
        service.spreadsheets()
        .get(
            spreadsheetId=SPREADSHEET_ID,
            ranges=[f"{ATTENDANCE_SHEET}!J{row_number}:AJ{row_number}"],
            includeGridData=True,
            fields=(
                "sheets(properties(sheetId,title),"
                "data(rowData(values(userEnteredValue,userEnteredFormat))))"
            ),
        )
        .execute()
    )
    row_data = (
        response.get("sheets", [{}])[0]
        .get("data", [{}])[0]
        .get("rowData", [{}])[0]
        .get("values", [])
    )

    cells: list[dict[str, Any]] = []
    for raw_cell in row_data[:ATTENDANCE_CELL_WIDTH]:
        cell: dict[str, Any] = {}
        if "userEnteredValue" in raw_cell:
            cell["userEnteredValue"] = copy.deepcopy(raw_cell["userEnteredValue"])
        if "userEnteredFormat" in raw_cell:
            cell["userEnteredFormat"] = copy.deepcopy(raw_cell["userEnteredFormat"])
        cells.append(cell)

    while len(cells) < ATTENDANCE_CELL_WIDTH:
        cells.append({})

    return cells


def attendance_cells_are_empty(cells: list[dict[str, Any]]) -> bool:
    for cell in cells:
        if normalize_extended_value(cell.get("userEnteredValue")) is not None:
            return False
    return True


def count_marked_attendance_cells(cells: list[dict[str, Any]]) -> int:
    return sum(
        1
        for cell in cells
        if normalize_extended_value(cell.get("userEnteredValue")) is not None
    )


def normalize_extended_value(value: dict[str, Any] | None) -> Any:
    if not value:
        return None

    if "stringValue" in value:
        text = str(value["stringValue"]).strip()
        if not text:
            return None
        return ("stringValue", text)

    if "numberValue" in value:
        return ("numberValue", float(value["numberValue"]))

    if "boolValue" in value:
        return ("boolValue", bool(value["boolValue"]))

    if "formulaValue" in value:
        return ("formulaValue", str(value["formulaValue"]))

    if "errorValue" in value:
        return ("errorValue", copy.deepcopy(value["errorValue"]))

    return ("raw", copy.deepcopy(value))


def normalize_cell_format(cell: dict[str, Any]) -> dict[str, Any]:
    return copy.deepcopy(cell.get("userEnteredFormat", {}))


def clear_cell_value(cell: dict[str, Any]) -> dict[str, Any]:
    cleared: dict[str, Any] = {}
    if "userEnteredFormat" in cell:
        cleared["userEnteredFormat"] = copy.deepcopy(cell["userEnteredFormat"])
    return cleared


def column_index_to_letter(column_index: int) -> str:
    result = ""
    current = column_index + 1
    while current > 0:
        current, remainder = divmod(current - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def build_attendance_merge_plan(
    source_cells: list[dict[str, Any]],
    destination_cells: list[dict[str, Any]],
    destination_row_number: int,
) -> AttendanceMergePlan:
    if len(source_cells) != len(destination_cells):
        raise PromotionError("Attendance rows have different widths.")

    merged_destination_cells: list[dict[str, Any]] = []
    cleared_source_cells: list[dict[str, Any]] = []

    for offset, (source_cell, destination_cell) in enumerate(zip(source_cells, destination_cells)):
        source_value = normalize_extended_value(source_cell.get("userEnteredValue"))
        destination_value = normalize_extended_value(destination_cell.get("userEnteredValue"))

        if source_value is None:
            merged_destination_cells.append(copy.deepcopy(destination_cell))
            cleared_source_cells.append(clear_cell_value(source_cell))
            continue

        if destination_value is None:
            merged_destination_cells.append(copy.deepcopy(source_cell))
            cleared_source_cells.append(clear_cell_value(source_cell))
            continue

        cell_ref = f"{column_index_to_letter(ATTENDANCE_START_COLUMN_INDEX + offset)}{destination_row_number}"
        if source_value != destination_value:
            raise PromotionError(
                f"Attendance target cell {cell_ref} conflicts with an existing value."
            )

        source_format = normalize_cell_format(source_cell)
        destination_format = normalize_cell_format(destination_cell)
        if source_format != destination_format:
            raise PromotionError(
                f"Attendance target cell {cell_ref} has the same value but a different format."
            )

        merged_destination_cells.append(copy.deepcopy(destination_cell))
        cleared_source_cells.append(clear_cell_value(source_cell))

    return AttendanceMergePlan(
        cleared_source_cells=cleared_source_cells,
        merged_destination_cells=merged_destination_cells,
    )


def build_grid_range(sheet_id: int, row_number: int) -> dict[str, int]:
    return {
        "sheetId": sheet_id,
        "startRowIndex": row_number - 1,
        "endRowIndex": row_number,
        "startColumnIndex": ATTENDANCE_START_COLUMN_INDEX,
        "endColumnIndex": ATTENDANCE_END_COLUMN_INDEX,
    }


def make_extended_value(value: Any) -> dict[str, Any]:
    if isinstance(value, bool):
        return {"boolValue": value}
    if isinstance(value, (int, float)):
        return {"numberValue": float(value)}
    return {"stringValue": str(value)}


def build_lineup_update_request(sheet_id: int, member: LineupMember, plan: PromotionPlan) -> dict[str, Any]:
    return {
        "updateCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": member.sheet_row - 1,
                "endRowIndex": member.sheet_row,
                "startColumnIndex": 0,
                "endColumnIndex": 4,
            },
            "rows": [
                {
                    "values": [
                        {"userEnteredValue": make_extended_value(plan.new_no)},
                        {"userEnteredValue": make_extended_value(plan.target_group.group_no)},
                        {"userEnteredValue": make_extended_value(plan.target_group.town)},
                        {"userEnteredValue": make_extended_value(plan.target_group.leader_name)},
                    ]
                }
            ],
            "fields": "userEnteredValue",
        }
    }


def build_attendance_row_update_request(
    sheet_id: int,
    row_number: int,
    cells: list[dict[str, Any]],
) -> dict[str, Any]:
    return {
        "updateCells": {
            "range": build_grid_range(sheet_id, row_number),
            "rows": [{"values": copy.deepcopy(cells)}],
            "fields": "userEnteredValue,userEnteredFormat",
        }
    }


def execute_promotion(
    service,
    metadata: dict[str, int],
    plan: PromotionPlan,
    attendance_merge_plan: AttendanceMergePlan,
) -> None:
    lineup_sheet_id = metadata.get(LINEUP_SHEET)
    attendance_sheet_id = metadata.get(ATTENDANCE_SHEET)
    if lineup_sheet_id is None or attendance_sheet_id is None:
        raise PromotionError("필수 시트 메타데이터를 찾지 못했습니다.")

    requests = [
        build_lineup_update_request(lineup_sheet_id, plan.source_member, plan),
        build_attendance_row_update_request(
            attendance_sheet_id,
            plan.new_attendance_row.sheet_row,
            attendance_merge_plan.merged_destination_cells,
        ),
        build_attendance_row_update_request(
            attendance_sheet_id,
            plan.old_attendance_row.sheet_row,
            attendance_merge_plan.cleared_source_cells,
        ),
    ]

    (
        service.spreadsheets()
        .batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": requests})
        .execute()
    )


def render_ebs_option(option: GroupOption) -> str:
    return f"{option.town} / {option.leader_name} / 조번호 {option.group_no} / 조원 {option.member_count}명"


def render_member_option(member: LineupMember) -> str:
    return f"{member.display_name} / No {member.no}"


def render_member_search_option(member: LineupMember) -> str:
    return (
        f"{member.display_name} / No {member.no} / "
        f"EBS {member.group_no} / {member.town} / {member.leader}"
    )


def render_town_option(town: str) -> str:
    return town


def render_gbs_option(option: GroupOption) -> str:
    return f"{option.town} / {option.leader_name} / 조번호 {option.group_no} / 인원 {option.member_count}명"


def clear_screen() -> None:
    os.system("cls" if os.name == "nt" else "clear")


def display_selection_menu(
    title: str,
    options: list[Any],
    current_index: int,
    render,
    escape_label: str = "취소",
) -> None:
    clear_screen()
    print("=" * 72)
    print(title)
    print("=" * 72)
    print(f"↑/↓: 이동, Enter: 선택, ESC: {escape_label}")
    print()

    for idx, option in enumerate(options):
        cursor = "▶ " if idx == current_index else "  "
        print(f"{cursor}{render(option)}")


def prompt_selection(title: str, options: list[Any], render, escape_label: str = "취소") -> Any:
    if not options:
        raise PromotionError(f"{title} 후보가 없습니다.")

    current_index = 0

    if os.name == "nt":
        import msvcrt

        while True:
            display_selection_menu(title, options, current_index, render, escape_label)
            key = msvcrt.getch()

            if key in {b"\xe0", b"\x00"}:
                key = msvcrt.getch()
                if key == b"H":
                    current_index = max(0, current_index - 1)
                elif key == b"P":
                    current_index = min(len(options) - 1, current_index + 1)
            elif key == b"\r":
                clear_screen()
                return options[current_index]
            elif key == b"\x1b":
                clear_screen()
                raise KeyboardInterrupt

    else:
        import termios
        import tty

        old_settings = termios.tcgetattr(sys.stdin)
        try:
            tty.setraw(sys.stdin.fileno())

            while True:
                display_selection_menu(title, options, current_index, render, escape_label)
                key = sys.stdin.read(1)

                if key == "\x1b":
                    next_key = sys.stdin.read(2)
                    if next_key == "[A":
                        current_index = max(0, current_index - 1)
                    elif next_key == "[B":
                        current_index = min(len(options) - 1, current_index + 1)
                    elif next_key == "":
                        clear_screen()
                        raise KeyboardInterrupt
                elif key == "\r":
                    clear_screen()
                    return options[current_index]

        finally:
            termios.tcsetattr(sys.stdin, termios.TCSADRAIN, old_settings)


def prompt_member_search(members: list[LineupMember]) -> LineupMember:
    searchable_members = build_searchable_ebs_members(members)
    if not searchable_members:
        raise PromotionError("EBS 범위에 검색 가능한 등반 대상자가 없습니다.")

    while True:
        clear_screen()
        print("=" * 72)
        print("1. EBS 내 등반 대상자 이름 검색")
        print("=" * 72)
        print("이름으로 검색한 뒤 결과에서 대상자를 선택합니다.")
        print("검색 범위는 EBS 일반 조원만 포함합니다. 빈값 입력 시 취소됩니다.")
        print()

        keyword = input("검색어: ").strip()
        if not keyword:
            raise KeyboardInterrupt

        member_options = search_ebs_members_by_name(searchable_members, keyword)
        if not member_options:
            print(f"\n'{keyword}' 검색 결과가 없습니다.")
            input("Enter를 누르면 다시 검색합니다...")
            continue

        try:
            return prompt_selection(
                f"1-1. 검색 결과에서 대상자 선택 ({keyword})",
                member_options,
                render_member_search_option,
                escape_label="검색으로 돌아가기",
            )
        except KeyboardInterrupt:
            continue


def print_plan_summary(plan: PromotionPlan, moved_marks: int) -> None:
    print()
    print("=" * 72)
    print("실행 미리보기")
    print("=" * 72)
    print(f"EBS: {plan.source_ebs.town} / {plan.source_ebs.leader_name} / 조번호 {plan.source_ebs.group_no}")
    print(f"대상자: {plan.source_member.display_name}")
    print(f"현재 No: {plan.old_no}")
    print(f"현재 소속: {plan.source_member.town} / {plan.source_member.leader} / 조번호 {plan.source_member.group_no}")
    print(f"이동 대상 GBS: {plan.target_group.town} / {plan.target_group.leader_name} / 조번호 {plan.target_group.group_no}")
    print(f"새 No: {plan.new_no}")
    print(
        "출석부 이동: "
        f"{plan.old_attendance_row.sheet_row}행 -> {plan.new_attendance_row.sheet_row}행 "
        f"(주간 출석 {moved_marks}칸)"
    )


def prompt_confirmation() -> bool:
    answer = input("\n실행할까요? [y/N]: ").strip().lower()
    return answer in {"y", "yes"}


def load_snapshot(service) -> tuple[dict[str, int], list[LineupMember], dict[int, AttendanceRowRef]]:
    metadata = fetch_sheet_metadata(service)
    lineup_rows, attendance_rows = fetch_values_snapshot(service)
    members = parse_lineup_members(lineup_rows)
    attendance_refs = parse_attendance_rows(attendance_rows)

    if not members:
        raise PromotionError("라인업정보에서 유효한 조원 데이터를 찾지 못했습니다.")
    if not attendance_refs:
        raise PromotionError("출석부에서 유효한 KEY 데이터를 찾지 못했습니다.")

    return metadata, members, attendance_refs


def validate_cli_args(argv: list[str]) -> None:
    if argv:
        joined = " ".join(argv)
        raise PromotionError(f"??? ??? ???? ????: {joined}")


def main(argv: list[str] | None = None) -> int:
    try:
        validate_cli_args(list(argv if argv is not None else sys.argv[1:]))

        service = build_service()
        metadata, members, attendance_refs = load_snapshot(service)

        ebs_options = build_ebs_options(members)
        selected_member = prompt_member_search(members)
        selected_ebs = find_ebs_option(ebs_options, selected_member.group_no)

        town_options = build_town_options(members)
        selected_town = prompt_selection("2. 이동할 GBS 고을 선택", town_options, render_town_option)

        gbs_options = build_gbs_options(members, selected_town)
        if not gbs_options:
            raise PromotionError("선택한 고을에 이동 가능한 GBS가 없습니다.")
        selected_gbs = prompt_selection("3. 이동할 GBS 선택", gbs_options, render_gbs_option)

        plan = build_promotion_plan(
            members,
            attendance_refs,
            selected_ebs,
            selected_member,
            selected_gbs,
        )

        source_cells = fetch_attendance_row_cells(service, plan.old_attendance_row.sheet_row)
        destination_cells = fetch_attendance_row_cells(service, plan.new_attendance_row.sheet_row)
        attendance_merge_plan = build_attendance_merge_plan(
            source_cells,
            destination_cells,
            plan.new_attendance_row.sheet_row,
        )

        moved_marks = count_marked_attendance_cells(source_cells)
        print_plan_summary(plan, moved_marks)

        if not prompt_confirmation():
            print("\n??? ??????.")
            return 0

        execute_promotion(service, metadata, plan, attendance_merge_plan)

        print()
        print("=" * 72)
        print("등반 처리 완료")
        print("=" * 72)
        print(f"{plan.source_member.name}: No {plan.old_no} -> {plan.new_no}")
        print(
            f"{plan.source_member.town}/{plan.source_member.leader} -> "
            f"{plan.target_group.town}/{plan.target_group.leader_name}"
        )
        print(f"출석부 주간 출석 이동: {moved_marks}칸")
        return 0

    except KeyboardInterrupt:
        print("\n작업을 취소했습니다.")
        return 0
    except PromotionError as exc:
        print(f"\n오류: {exc}")
        return 1
    except Exception as exc:
        print(f"\n예상하지 못한 오류: {exc}")
        import traceback

        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())

