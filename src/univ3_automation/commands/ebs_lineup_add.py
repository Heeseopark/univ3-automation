#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EBS 라인업 일괄 추가 스크립트.

선택한 EBS에 조원 텍스트를 붙여넣으면 파싱/정렬 후 라인업정보 시트에 행을 추가한다.
"""

from __future__ import annotations

try:
    from ebs_to_gbs_promotion import (
        LINEUP_SHEET,
        SPREADSHEET_ID,
        GroupOption,
        LineupMember,
        PromotionError,
        EBS_GROUP_NUMBERS,
        build_service,
        cell_value,
        fetch_sheet_metadata,
        make_extended_value,
        parse_int,
        parse_lineup_members,
        prompt_confirmation,
        prompt_selection,
        render_ebs_option,
    )
except ImportError:
    from univ3_automation.commands.ebs_to_gbs_promotion import (
        LINEUP_SHEET,
        SPREADSHEET_ID,
        GroupOption,
        LineupMember,
        PromotionError,
        EBS_GROUP_NUMBERS,
        build_service,
        cell_value,
        fetch_sheet_metadata,
        make_extended_value,
        parse_int,
        parse_lineup_members,
        prompt_confirmation,
        prompt_selection,
        render_ebs_option,
    )

import copy
import re
import sys
from dataclasses import dataclass
from typing import Any

LINEUP_END_COLUMN = "AO"
ROW_CELL_WIDTH = 41
GRADE_SORT_FALLBACK = -1
DEFAULT_BACKGROUND = {"red": 1.0, "green": 1.0, "blue": 1.0}
YELLOW_BACKGROUND = {"red": 1.0, "green": 1.0, "blue": 0.0}
H_COLUMN_INDEX = 7
I_COLUMN_INDEX = 8
J_COLUMN_INDEX = 9
O_COLUMN_INDEX = 14
AM_COLUMN_INDEX = 38
AN_COLUMN_INDEX = 39
AO_COLUMN_INDEX = 40
HIGHLIGHT_COLUMN_INDEXES = (O_COLUMN_INDEX, AM_COLUMN_INDEX, AN_COLUMN_INDEX, AO_COLUMN_INDEX)
ENTRY_PATTERN = re.compile(r"^(남|여)\s*(\d+)\s+(.+)$")


@dataclass(frozen=True)
class ParsedMember:
    gender: str
    grade: int
    name: str
    no: int | None = None
    is_new: bool = True

    @property
    def display_name(self) -> str:
        return f"{self.gender}{self.grade} {self.name}"


@dataclass(frozen=True)
class ExistingMemberRow:
    member: LineupMember
    cells: list[dict[str, Any]]


@dataclass(frozen=True)
class GroupBlock:
    root_row: int
    end_row: int
    general_row_numbers: list[int]
    template_row: int

    @property
    def insert_row(self) -> int:
        return self.end_row + 1


@dataclass(frozen=True)
class SortedAssignment:
    target_row: int
    member_no: int
    gender: str
    grade: int
    name: str
    source_cells: list[dict[str, Any]]
    is_new: bool

    @property
    def display_name(self) -> str:
        return f"{self.gender}{self.grade} {self.name}"


def fetch_lineup_values(service) -> list[list[Any]]:
    response = (
        service.spreadsheets()
        .values()
        .get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{LINEUP_SHEET}!A1:{LINEUP_END_COLUMN}",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
    )
    rows = response.get("values", [])
    if not rows:
        raise PromotionError("라인업정보 시트가 비어 있습니다.")
    return rows


def fetch_lineup_block_cells(
    service,
    start_row: int,
    end_row: int,
) -> dict[int, list[dict[str, Any]]]:
    response = (
        service.spreadsheets()
        .get(
            spreadsheetId=SPREADSHEET_ID,
            ranges=[f"{LINEUP_SHEET}!A{start_row}:{LINEUP_END_COLUMN}{end_row}"],
            includeGridData=True,
            fields="sheets(data(rowData(values(userEnteredValue,userEnteredFormat))))",
        )
        .execute()
    )
    row_data = response.get("sheets", [{}])[0].get("data", [{}])[0].get("rowData", [])
    result: dict[int, list[dict[str, Any]]] = {}
    for offset, raw_row in enumerate(row_data):
        row_number = start_row + offset
        values = raw_row.get("values", [])
        cells: list[dict[str, Any]] = []
        for raw_cell in values[:ROW_CELL_WIDTH]:
            cell: dict[str, Any] = {}
            if "userEnteredValue" in raw_cell:
                cell["userEnteredValue"] = copy.deepcopy(raw_cell["userEnteredValue"])
            if "userEnteredFormat" in raw_cell:
                cell["userEnteredFormat"] = copy.deepcopy(raw_cell["userEnteredFormat"])
            cells.append(cell)
        while len(cells) < ROW_CELL_WIDTH:
            cells.append({})
        result[row_number] = cells
    return result


def parse_member_text(raw_text: str) -> list[ParsedMember]:
    tokens = [token.strip() for token in re.split(r"[\n,]+", raw_text) if token.strip()]
    if not tokens:
        raise PromotionError("추가할 조원 텍스트가 비어 있습니다.")

    members: list[ParsedMember] = []
    invalid_tokens: list[str] = []
    for token in tokens:
        match = ENTRY_PATTERN.match(token)
        if match is None:
            invalid_tokens.append(token)
            continue
        gender, grade_text, name = match.groups()
        members.append(ParsedMember(gender=gender, grade=int(grade_text), name=name.strip()))

    if invalid_tokens:
        joined = ", ".join(invalid_tokens)
        raise PromotionError(f"파싱할 수 없는 입력이 있습니다: {joined}")

    return members


def sort_key(grade: int, name: str) -> tuple[int, str]:
    return (-grade, name)


def sort_parsed_members(members: list[ParsedMember]) -> list[ParsedMember]:
    return sorted(members, key=lambda item: sort_key(item.grade, item.name))


def select_group_root_member(group_members: list[LineupMember], group_no: int) -> LineupMember:
    return min(
        group_members,
        key=lambda member: (
            0 if member.no == group_no * 100 + 1 else 1,
            member.sheet_row,
            member.no,
        ),
    )


def build_lineup_ebs_options(members: list[LineupMember]) -> list[GroupOption]:
    grouped: dict[int, list[LineupMember]] = {}
    for member in members:
        grouped.setdefault(member.group_no, []).append(member)

    options: list[GroupOption] = []
    for group_no in EBS_GROUP_NUMBERS:
        group_members = grouped.get(group_no)
        if not group_members:
            continue
        root = select_group_root_member(group_members, group_no)
        member_count = sum(1 for row in group_members if row.is_movable)
        options.append(
            GroupOption(
                group_no=group_no,
                town=root.town,
                leader_name=root.name,
                leader_role=root.leader_role,
                member_count=member_count,
            )
        )
    return options


def assign_new_member_numbers(members: list[ParsedMember], starting_no: int) -> list[ParsedMember]:
    assigned: list[ParsedMember] = []
    for offset, member in enumerate(members):
        assigned.append(
            ParsedMember(
                gender=member.gender,
                grade=member.grade,
                name=member.name,
                no=starting_no + offset,
                is_new=True,
            )
        )
    return assigned


def build_group_block(rows: list[list[Any]], members: list[LineupMember], group_no: int) -> GroupBlock:
    group_members = [member for member in members if member.group_no == group_no]
    if not group_members:
        raise PromotionError("선택한 EBS 데이터를 찾지 못했습니다.")

    root = select_group_root_member(group_members, group_no)
    end_row = root.sheet_row
    for row_number in range(root.sheet_row, len(rows) + 1):
        row = rows[row_number - 1]
        if parse_int(cell_value(row, 1)) != group_no:
            break
        end_row = row_number

    general_row_numbers = sorted(
        member.sheet_row
        for member in group_members
        if root.sheet_row <= member.sheet_row <= end_row and member.is_movable
    )
    template_row = general_row_numbers[-1] if general_row_numbers else root.sheet_row
    return GroupBlock(
        root_row=root.sheet_row,
        end_row=end_row,
        general_row_numbers=general_row_numbers,
        template_row=template_row,
    )


def grade_to_int(value: str) -> int:
    parsed = parse_int(value)
    if parsed is None:
        return GRADE_SORT_FALLBACK
    return parsed


def build_blank_row_from_template(template_cells: list[dict[str, Any]]) -> list[dict[str, Any]]:
    blank_row: list[dict[str, Any]] = []
    for template_cell in template_cells[:ROW_CELL_WIDTH]:
        user_format = copy.deepcopy(template_cell.get("userEnteredFormat", {}))
        user_format.pop("backgroundColor", None)
        user_format.pop("backgroundColorStyle", None)
        user_format["backgroundColor"] = copy.deepcopy(DEFAULT_BACKGROUND)
        blank_row.append({"userEnteredFormat": user_format})
    while len(blank_row) < ROW_CELL_WIDTH:
        blank_row.append({"userEnteredFormat": {"backgroundColor": copy.deepcopy(DEFAULT_BACKGROUND)}})
    return blank_row


def set_cell_value(cells: list[dict[str, Any]], index: int, value: Any) -> None:
    if value is None:
        cells[index].pop("userEnteredValue", None)
        return
    cells[index]["userEnteredValue"] = make_extended_value(value)


def set_cell_formula(cells: list[dict[str, Any]], index: int, formula: str) -> None:
    cells[index]["userEnteredValue"] = {"formulaValue": formula}


def set_cell_background(cells: list[dict[str, Any]], index: int, background_color: dict[str, float]) -> None:
    user_format = copy.deepcopy(cells[index].get("userEnteredFormat", {}))
    user_format["backgroundColor"] = copy.deepcopy(background_color)
    cells[index]["userEnteredFormat"] = user_format


def build_attendance_formula(row_number: int) -> str:
    return f'=IFERROR(VLOOKUP(A{row_number},\'출석부\'!$D$7:$H$1118,5,FALSE), " ")'


def build_grade_formula(row_number: int) -> str:
    return (
        f'=IF(H{row_number}=" "," ",'
        f'IF(H{row_number}>=0.8,"A",'
        f'IF(H{row_number}>=0.5,"B",'
        f'IF(H{row_number}>=0.25,"C","D"))))'
    )


def build_existing_member_cells(existing: ExistingMemberRow, target_row: int) -> list[dict[str, Any]]:
    cells = copy.deepcopy(existing.cells)
    member = existing.member
    set_cell_value(cells, 0, member.no)
    set_cell_value(cells, 1, member.group_no)
    set_cell_value(cells, 2, member.town)
    set_cell_value(cells, 3, member.leader)
    set_cell_value(cells, 4, member.gender)
    set_cell_value(cells, 5, grade_to_int(member.grade))
    set_cell_value(cells, 6, member.name)
    set_cell_formula(cells, H_COLUMN_INDEX, build_attendance_formula(target_row))
    set_cell_formula(cells, I_COLUMN_INDEX, build_grade_formula(target_row))
    return cells


def build_new_member_cells(
    template_cells: list[dict[str, Any]],
    source_ebs: GroupOption,
    member: ParsedMember,
    target_row: int,
) -> list[dict[str, Any]]:
    if member.no is None:
        raise PromotionError("신규 조원의 No가 할당되지 않았습니다.")

    cells = build_blank_row_from_template(template_cells)
    set_cell_value(cells, 0, member.no)
    set_cell_value(cells, 1, source_ebs.group_no)
    set_cell_value(cells, 2, source_ebs.town)
    set_cell_value(cells, 3, source_ebs.leader_name)
    set_cell_value(cells, 4, member.gender)
    set_cell_value(cells, 5, member.grade)
    set_cell_value(cells, 6, member.name)
    set_cell_formula(cells, H_COLUMN_INDEX, build_attendance_formula(target_row))
    set_cell_formula(cells, I_COLUMN_INDEX, build_grade_formula(target_row))
    set_cell_value(cells, J_COLUMN_INDEX, None)

    for column_index in HIGHLIGHT_COLUMN_INDEXES:
        set_cell_value(cells, column_index, None)
        set_cell_background(cells, column_index, YELLOW_BACKGROUND)

    return cells


def build_sorted_assignments(
    existing_rows: list[ExistingMemberRow],
    new_members: list[ParsedMember],
    target_rows: list[int],
    template_cells: list[dict[str, Any]],
    source_ebs: GroupOption,
) -> list[SortedAssignment]:
    sortable: list[tuple[tuple[int, str], SortedAssignment]] = []
    for existing in existing_rows:
        sortable.append(
            (
                sort_key(grade_to_int(existing.member.grade), existing.member.name),
                SortedAssignment(
                    target_row=0,
                    member_no=existing.member.no,
                    gender=existing.member.gender,
                    grade=grade_to_int(existing.member.grade),
                    name=existing.member.name,
                    source_cells=build_existing_member_cells(existing, existing.member.sheet_row),
                    is_new=False,
                ),
            )
        )
    for member in new_members:
        sortable.append(
            (
                sort_key(member.grade, member.name),
                SortedAssignment(
                    target_row=0,
                    member_no=member.no or 0,
                    gender=member.gender,
                    grade=member.grade,
                    name=member.name,
                    source_cells=build_new_member_cells(template_cells, source_ebs, member, 0),
                    is_new=True,
                ),
            )
        )

    sortable.sort(key=lambda item: item[0])
    if len(sortable) != len(target_rows):
        raise PromotionError("정렬 대상 수와 배치 대상 row 수가 일치하지 않습니다.")

    assignments: list[SortedAssignment] = []
    for target_row, (_, assignment) in zip(target_rows, sortable):
        cells = copy.deepcopy(assignment.source_cells)
        set_cell_formula(cells, H_COLUMN_INDEX, build_attendance_formula(target_row))
        set_cell_formula(cells, I_COLUMN_INDEX, build_grade_formula(target_row))
        assignments.append(
            SortedAssignment(
                target_row=target_row,
                member_no=assignment.member_no,
                gender=assignment.gender,
                grade=assignment.grade,
                name=assignment.name,
                source_cells=cells,
                is_new=assignment.is_new,
            )
        )
    return assignments


def build_insert_rows_request(sheet_id: int, start_row: int, row_count: int) -> dict[str, Any]:
    return {
        "insertDimension": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "ROWS",
                "startIndex": start_row - 1,
                "endIndex": start_row - 1 + row_count,
            },
            "inheritFromBefore": False,
        }
    }


def build_row_update_request(sheet_id: int, row_number: int, cells: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "updateCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": row_number - 1,
                "endRowIndex": row_number,
                "startColumnIndex": 0,
                "endColumnIndex": ROW_CELL_WIDTH,
            },
            "rows": [{"values": copy.deepcopy(cells)}],
            "fields": "userEnteredValue,userEnteredFormat",
        }
    }


def prompt_multiline_text() -> str:
    print("=" * 72)
    print("2. 추가할 EBS 조원 텍스트 입력")
    print("=" * 72)
    print("예시: 여4 이지윤, 여3 권호진, 남1 문준혁")
    print("붙여넣은 뒤 빈 줄에서 Enter를 누르면 입력이 끝납니다.")

    lines: list[str] = []
    while True:
        line = input("> " if not lines else "")
        if not line.strip():
            break
        lines.append(line.rstrip())

    text = "\n".join(lines).strip()
    if not text:
        raise PromotionError("추가할 조원 텍스트가 입력되지 않았습니다.")
    return text


def print_plan_summary(
    source_ebs: GroupOption,
    block: GroupBlock,
    starting_no: int,
    assignments: list[SortedAssignment],
) -> None:
    print()
    print("=" * 72)
    print("실행 미리보기")
    print("=" * 72)
    print(f"EBS: {source_ebs.town} / {source_ebs.leader_name} / 조번호 {source_ebs.group_no}")
    print(f"삽입 시작 row: {block.insert_row}")
    print(f"신규 시작 No: {starting_no}")
    print()
    for assignment in assignments:
        label = "[NEW]" if assignment.is_new else "[MOVE]"
        print(f"{label} row {assignment.target_row} / No {assignment.member_no} / {assignment.display_name}")


def load_snapshot(service) -> tuple[dict[str, int], list[list[Any]], list[LineupMember]]:
    metadata = fetch_sheet_metadata(service)
    lineup_rows = fetch_lineup_values(service)
    members = parse_lineup_members(lineup_rows)

    if not members:
        raise PromotionError("라인업정보에서 유효한 조원 데이터를 찾지 못했습니다.")
    return metadata, lineup_rows, members


def validate_cli_args(argv: list[str]) -> None:
    if argv:
        joined = " ".join(argv)
        raise PromotionError(f"지원하지 않는 옵션입니다: {joined}")


def build_existing_member_rows(
    members: list[LineupMember],
    block: GroupBlock,
    row_cells: dict[int, list[dict[str, Any]]],
) -> list[ExistingMemberRow]:
    existing: list[ExistingMemberRow] = []
    members_by_row = {member.sheet_row: member for member in members}
    for row_number in block.general_row_numbers:
        member = members_by_row.get(row_number)
        if member is None:
            raise PromotionError(f"{row_number}행의 조원 데이터를 찾지 못했습니다.")
        cells = row_cells.get(row_number)
        if cells is None:
            raise PromotionError(f"{row_number}행의 셀 데이터를 읽지 못했습니다.")
        existing.append(ExistingMemberRow(member=member, cells=cells))
    return existing


def next_member_no(members: list[LineupMember], group_no: int) -> int:
    group_member_nos = [member.no for member in members if member.group_no == group_no]
    if not group_member_nos:
        raise PromotionError("선택한 EBS의 No 정보를 찾지 못했습니다.")
    return max(group_member_nos) + 1


def execute_lineup_add(
    service,
    metadata: dict[str, int],
    insert_row: int,
    insert_count: int,
    assignments: list[SortedAssignment],
) -> None:
    lineup_sheet_id = metadata.get(LINEUP_SHEET)
    if lineup_sheet_id is None:
        raise PromotionError("라인업정보 시트 메타데이터를 찾지 못했습니다.")

    requests = [build_insert_rows_request(lineup_sheet_id, insert_row, insert_count)]
    for assignment in assignments:
        requests.append(build_row_update_request(lineup_sheet_id, assignment.target_row, assignment.source_cells))

    (
        service.spreadsheets()
        .batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": requests})
        .execute()
    )


def main(argv: list[str] | None = None) -> int:
    try:
        validate_cli_args(list(argv if argv is not None else sys.argv[1:]))

        service = build_service()
        metadata, lineup_rows, members = load_snapshot(service)

        ebs_options = build_lineup_ebs_options(members)
        selected_ebs = prompt_selection("1. EBS 선택", ebs_options, render_ebs_option)

        raw_text = prompt_multiline_text()
        parsed_members = sort_parsed_members(parse_member_text(raw_text))
        numbered_new_members = assign_new_member_numbers(
            parsed_members,
            next_member_no(members, selected_ebs.group_no),
        )

        block = build_group_block(lineup_rows, members, selected_ebs.group_no)
        row_cells = fetch_lineup_block_cells(service, block.root_row, block.end_row)
        template_cells = row_cells.get(block.template_row)
        if template_cells is None:
            raise PromotionError("새 row의 서식 기준 행을 읽지 못했습니다.")

        group_members = [member for member in members if member.group_no == selected_ebs.group_no]
        existing_member_rows = build_existing_member_rows(group_members, block, row_cells)
        target_rows = block.general_row_numbers + list(range(block.insert_row, block.insert_row + len(numbered_new_members)))
        assignments = build_sorted_assignments(
            existing_member_rows,
            numbered_new_members,
            target_rows,
            template_cells,
            selected_ebs,
        )

        print_plan_summary(selected_ebs, block, numbered_new_members[0].no or 0, assignments)

        if not prompt_confirmation():
            print("\n작업을 취소했습니다.")
            return 0

        execute_lineup_add(
            service,
            metadata,
            block.insert_row,
            len(numbered_new_members),
            assignments,
        )

        print()
        print("=" * 72)
        print("EBS 라인업 추가 완료")
        print("=" * 72)
        print(f"추가 인원: {len(numbered_new_members)}명")
        print(f"신규 No 범위: {numbered_new_members[0].no} - {numbered_new_members[-1].no}")
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




