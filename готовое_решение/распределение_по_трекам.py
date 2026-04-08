from __future__ import annotations

import math
import re
import zipfile
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from statistics import mean
from typing import Dict, List, Optional, Sequence, Tuple
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

OUTPUT_FOLDER_NAME = "выход"
SAMPLE_INPUT_FILE_NAME = "образец_входных_данных.xlsx"
OUTPUT_FILE_NAME = "результат_распределения.xlsx"

SEMESTER_AVERAGE_SHARE = 0.25
OVERALL_AVERAGE_SHARE = 0.15
SUBJECT_GRADES_SHARE = 0.60
PROGRESS_SIGNIFICANCE = 4.0
ATTENDANCE_SIGNIFICANCE = 4.0
EVENTS_SIGNIFICANCE = 2.0
OLYMPIADS_SIGNIFICANCE = 3.0
CONTESTS_SIGNIFICANCE = 2.0

MAIN_NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
PKG_NS = {"p": "http://schemas.openxmlformats.org/package/2006/relationships"}

OPTIONAL_COLUMN_ALIASES = {
    "посещаемость": "attendance_percent",
    "посещаемость %": "attendance_percent",
    "процент посещаемости": "attendance_percent",
    "внеучебные мероприятия": "events_count",
    "мероприятия": "events_count",
    "внеучебная активность": "events_count",
    "олимпиады": "olympiads_count",
    "конкурсы": "contests_count",
    "статус без выбора": "no_choice_status",
    "статус выбора трека": "no_choice_status",
    "ситуация без выбора": "no_choice_status",
}


def normalize_text(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\xa0", " ").strip().lower().replace("ё", "е")
    return re.sub(r"\s+", " ", text)


def normalize_subject_name(value: str) -> str:
    text = normalize_text(value)
    text = re.sub(r"\([^)]*\)", "", text)
    text = re.sub(r"[^a-zа-я0-9 ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return {
        "линейная алгебра и аналитическая геометрия": "линейная алгебра",
        "бжд обзр": "бжд",
        "объектно ориентированное программирование": "ооп",
    }.get(text, text)


def parse_float(value: object) -> Optional[float]:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    text = text.replace("%", "").replace(" ", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def parse_int(value: object) -> Optional[int]:
    number = parse_float(value)
    return None if number is None else int(round(number))


def safe_mean(values: Sequence[float], default: float = 0.0) -> float:
    return float(mean(values)) if values else default


def column_letters_to_index(letters: str) -> int:
    value = 0
    for letter in letters:
        value = value * 26 + (ord(letter.upper()) - ord("A") + 1)
    return value


def index_to_column_letters(index: int) -> str:
    letters: List[str] = []
    while index > 0:
        index, rest = divmod(index - 1, 26)
        letters.append(chr(ord("A") + rest))
    return "".join(reversed(letters))


def parse_cell_ref(ref: str) -> Tuple[int, int]:
    match = re.match(r"([A-Z]+)(\d+)", ref)
    if not match:
        raise ValueError(f"Некорректная ссылка на ячейку: {ref}")
    letters, row = match.groups()
    return int(row), column_letters_to_index(letters)


def extract_semester_number(value: str) -> Optional[int]:
    match = re.search(r"(\d+)\s*сем", normalize_text(value))
    return int(match.group(1)) if match else None


def extract_olympiad_value_index(header: str) -> Optional[int]:
    normalized = normalize_text(header)
    patterns = [
        r"олимпиада\s*№?\s*(\d+)$",
        r"олимпиада_(\d+)$",
    ]
    for pattern in patterns:
        match = re.fullmatch(pattern, normalized)
        if match:
            return int(match.group(1))
    return None


def extract_olympiad_scale_index(header: str) -> Optional[int]:
    normalized = normalize_text(header)
    patterns = [
        r"(?:вес|скейл|scale)\s*олимпиады?\s*№?\s*(\d+)$",
        r"олимпиада\s*№?\s*(\d+)\s*(?:вес|скейл|scale)$",
    ]
    for pattern in patterns:
        match = re.fullmatch(pattern, normalized)
        if match:
            return int(match.group(1))
    return None


def is_reserved_header(header: str) -> bool:
    normalized = normalize_text(header)
    return (
        not normalized
        or normalized == "id"
        or normalized.startswith("трек")
        or normalized in {"средний балл", "сумма баллов", "рейтинг", "место в рейтинге"}
        or normalized in OPTIONAL_COLUMN_ALIASES
        or extract_olympiad_value_index(header) is not None
        or extract_olympiad_scale_index(header) is not None
    )


class WorkbookData:
    def __init__(self, sheets: List[dict]) -> None:
        self.sheets = sheets


def read_xml_from_zip(archive: zipfile.ZipFile, name: str) -> Optional[ET.Element]:
    try:
        with archive.open(name) as handle:
            return ET.fromstring(handle.read())
    except KeyError:
        return None


def read_xlsx(path: Path) -> WorkbookData:
    with zipfile.ZipFile(path, "r") as archive:
        shared_strings_root = read_xml_from_zip(archive, "xl/sharedStrings.xml")
        shared_strings: List[str] = []
        if shared_strings_root is not None:
            for si in shared_strings_root.findall("m:si", MAIN_NS):
                shared_strings.append("".join((node.text or "") for node in si.findall(".//m:t", MAIN_NS)))

        workbook_root = read_xml_from_zip(archive, "xl/workbook.xml")
        rels_root = read_xml_from_zip(archive, "xl/_rels/workbook.xml.rels")
        if workbook_root is None or rels_root is None:
            raise ValueError("Не удалось прочитать файл Excel.")

        rel_map: Dict[str, str] = {}
        for rel in rels_root.findall("p:Relationship", PKG_NS):
            if rel.attrib.get("Id") and rel.attrib.get("Target"):
                rel_map[rel.attrib["Id"]] = rel.attrib["Target"]

        sheets: List[dict] = []
        for sheet in workbook_root.findall("m:sheets/m:sheet", MAIN_NS):
            rel_id = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            target = rel_map.get(rel_id or "")
            if not target:
                continue
            sheet_path = target if target.startswith("xl/") else f"xl/{target}"
            sheet_root = read_xml_from_zip(archive, sheet_path)
            if sheet_root is None:
                continue
            rows: Dict[int, Dict[int, str]] = defaultdict(dict)
            max_row = 0
            max_col = 0
            for cell in sheet_root.findall(".//m:sheetData/m:row/m:c", MAIN_NS):
                ref = cell.attrib.get("r")
                if not ref:
                    continue
                row_number, col_number = parse_cell_ref(ref)
                max_row = max(max_row, row_number)
                max_col = max(max_col, col_number)
                cell_type = cell.attrib.get("t")
                value_node = cell.find("m:v", MAIN_NS)
                inline_node = cell.find("m:is", MAIN_NS)
                value = ""
                if cell_type == "s" and value_node is not None and value_node.text is not None:
                    index = int(float(value_node.text))
                    value = shared_strings[index] if 0 <= index < len(shared_strings) else ""
                elif cell_type == "inlineStr" and inline_node is not None:
                    value = "".join((node.text or "") for node in inline_node.findall(".//m:t", MAIN_NS))
                elif value_node is not None and value_node.text is not None:
                    value = value_node.text
                rows[row_number][col_number] = value
            sheets.append({"title": sheet.attrib.get("name", "Лист1"), "rows": dict(rows), "max_row": max_row, "max_col": max_col})
    return WorkbookData(sheets)


def get_cell(sheet: dict, row_number: int, col_number: int) -> str:
    return sheet["rows"].get(row_number, {}).get(col_number, "")


def find_input_sheet(workbook: WorkbookData) -> Tuple[dict, int]:
    fallback = None
    for sheet in workbook.sheets:
        for row_number in range(1, min(sheet["max_row"], 12) + 1):
            row_values = [normalize_text(get_cell(sheet, row_number, col)) for col in range(1, sheet["max_col"] + 1)]
            has_id = "id" in row_values
            track_headers = [value for value in row_values if value.startswith("трек")]
            if has_id and len(track_headers) >= 2:
                return sheet, row_number
            if has_id and fallback is None:
                fallback = (sheet, row_number)
    if fallback is None:
        raise ValueError("Не найден лист с колонкой id и треками.")
    return fallback


def read_header_map(sheet: dict, header_row: int) -> Dict[int, str]:
    headers: Dict[int, str] = {}
    for col in range(1, sheet["max_col"] + 1):
        value = get_cell(sheet, header_row, col).strip()
        if value:
            headers[col] = value
    return headers


def read_semester_marks(sheet: dict, semester_row: int, max_col: int) -> Dict[int, Optional[int]]:
    marks: Dict[int, Optional[int]] = {}
    current_semester: Optional[int] = None
    for col in range(1, max_col + 1):
        value = get_cell(sheet, semester_row, col).strip()
        number = extract_semester_number(value) if value else None
        if number is not None:
            current_semester = number
        marks[col] = current_semester
    return marks


def detect_optional_columns(headers: Dict[int, str]) -> Dict[str, int]:
    columns: Dict[str, int] = {}
    for col, header in headers.items():
        normalized = normalize_text(header)
        if normalized in OPTIONAL_COLUMN_ALIASES:
            columns[OPTIONAL_COLUMN_ALIASES[normalized]] = col
    return columns


def detect_olympiad_columns(headers: Dict[int, str]) -> List[dict]:
    value_columns: Dict[int, int] = {}
    scale_columns: Dict[int, int] = {}
    for col, header in headers.items():
        value_index = extract_olympiad_value_index(header)
        if value_index is not None:
            value_columns[value_index] = col
        scale_index = extract_olympiad_scale_index(header)
        if scale_index is not None:
            scale_columns[scale_index] = col
    slots: List[dict] = []
    for slot_index in sorted(value_columns):
        slots.append(
            {
                "index": slot_index,
                "value_col": value_columns[slot_index],
                "scale_col": scale_columns.get(slot_index),
            }
        )
    return slots


def calculate_olympiad_score(sheet: dict, row: int, olympiad_columns: Sequence[dict]) -> float:
    total = 0.0
    for slot in olympiad_columns:
        raw_value = get_cell(sheet, row, slot["value_col"]).strip()
        raw_scale = get_cell(sheet, row, slot["scale_col"]).strip() if slot["scale_col"] is not None else ""
        value = parse_float(raw_value)
        scale = parse_float(raw_scale)
        if value is None and scale is None:
            continue
        safe_value = 0.0 if value is None else value
        safe_scale = 1.0 if scale is None else scale
        total += safe_value * safe_scale
    return total


def extract_track_columns(headers: Dict[int, str]) -> List[Tuple[str, int]]:
    tracks = [(header.strip(), col) for col, header in headers.items() if normalize_text(header).startswith("трек")]
    if len(tracks) < 2:
        raise ValueError("Нужно минимум два столбца с приоритетами треков.")
    return tracks


def extract_grade_columns(headers: Dict[int, str], semester_marks: Dict[int, Optional[int]]) -> List[dict]:
    grade_columns: List[dict] = []
    for col, header in headers.items():
        semester_number = semester_marks.get(col)
        if semester_number is None or is_reserved_header(header):
            continue
        grade_columns.append(
            {
                "col": col,
                "header": header,
                "semester": semester_number,
                "normalized_subject": normalize_subject_name(header),
            }
        )
    if not grade_columns:
        raise ValueError("Колонки с оценками не найдены.")
    return grade_columns


def compute_semester_weights(semesters: Sequence[int]) -> Dict[int, float]:
    ordered = sorted(set(semesters))
    denominator = sum(range(1, len(ordered) + 1))
    return {semester: index / denominator for index, semester in enumerate(ordered, start=1)}


def normalize_track_preferences(track_values: Dict[str, Optional[int]], track_names: Sequence[str]) -> Dict[str, int]:
    result: Dict[str, int] = {}
    for track in track_names:
        value = track_values.get(track)
        result[track] = value if value is not None and value > 0 else 0
    return result


def choose_track(preferences: Dict[str, int], track_names: Sequence[str], loads: Dict[str, int], capacities: Dict[str, int]) -> str:
    available = [track for track in track_names if loads[track] < capacities[track]]
    if not available:
        raise ValueError("Не осталось доступных треков.")
    order = {name: index for index, name in enumerate(track_names)}
    positive_ranks = sorted({rank for rank in preferences.values() if rank > 0})
    for rank in positive_ranks:
        candidates = [track for track in available if preferences.get(track, 0) == rank]
        if candidates:
            return min(candidates, key=lambda track: (loads[track], order[track]))
    indifferent = [track for track in available if preferences.get(track, 0) == 0]
    if indifferent:
        return min(indifferent, key=lambda track: (loads[track], order[track]))
    return min(available, key=lambda track: (loads[track], order[track]))


def build_balanced_capacities(track_names: Sequence[str], student_count: int) -> Dict[str, int]:
    base = student_count // len(track_names)
    extra = student_count % len(track_names)
    return {track_name: base + (1 if index < extra else 0) for index, track_name in enumerate(track_names)}


def build_subject_order(students: Sequence[dict]) -> Dict[str, int]:
    subjects = sorted({item["normalized_subject"] for student in students for item in student["grades"]})
    return {subject: index + 1 for index, subject in enumerate(subjects)}


def calculate_id_micro_bonus(student_id: str) -> float:
    digits = "".join(symbol for symbol in str(student_id) if symbol.isdigit())
    if digits:
        digit_signature = sum((index + 1) * int(digit) for index, digit in enumerate(reversed(digits[-12:])))
    else:
        digit_signature = sum((index + 1) * ord(symbol) for index, symbol in enumerate(str(student_id)))
    return (digit_signature % 1_000) / 1_000_000.0


def calculate_student_rating(
    student: dict,
    semester_weights: Dict[int, float],
    cohort_maxima: dict,
    subject_order: Dict[str, int],
) -> None:
    grades_by_semester: Dict[int, List[float]] = defaultdict(list)
    subject_history: Dict[str, List[Tuple[int, float]]] = defaultdict(list)
    all_grades: List[float] = []
    for item in student["grades"]:
        grades_by_semester[item["semester"]].append(item["grade"])
        subject_history[item["normalized_subject"]].append((item["semester"], item["grade"]))
        all_grades.append(item["grade"])

    semester_averages = {semester: safe_mean(values) for semester, values in grades_by_semester.items()}
    subject_num = 0.0
    subject_den = 0.0
    for item in student["grades"]:
        weight = semester_weights[item["semester"]]
        subject_num += item["grade"] * weight
        subject_den += weight
    weighted_subject_average = subject_num / subject_den if subject_den else 0.0

    semester_num = 0.0
    semester_den = 0.0
    for semester, avg_grade in semester_averages.items():
        weight = semester_weights[semester]
        semester_num += avg_grade * weight
        semester_den += weight
    weighted_semester_average = semester_num / semester_den if semester_den else 0.0
    overall_average = safe_mean(all_grades)

    progress_sum = 0.0
    progress_count = 0
    for history in subject_history.values():
        if len(history) < 2:
            continue
        ordered = sorted(history, key=lambda item: item[0])
        for previous, current in zip(ordered, ordered[1:]):
            progress_sum += (current[1] - previous[1]) * semester_weights[current[0]]
            progress_count += 1
    progress_bonus = PROGRESS_SIGNIFICANCE * (progress_sum / progress_count if progress_count else 0.0)

    attendance_bonus = 0.0
    attendance = student["optional"].get("attendance_percent")
    if attendance is not None:
        attendance_bonus = max(0.0, min(attendance, 100.0)) / 100.0 * ATTENDANCE_SIGNIFICANCE

    events_bonus = 0.0
    events = student["optional"].get("events_count")
    if events is not None and cohort_maxima["events_count"] > 0:
        events_bonus = events / cohort_maxima["events_count"] * EVENTS_SIGNIFICANCE

    olympiads_bonus = 0.0
    olympiad_score = student.get("olympiad_score", 0.0)
    if olympiad_score > 0 and cohort_maxima["olympiad_score"] > 0:
        olympiads_bonus = olympiad_score / cohort_maxima["olympiad_score"] * OLYMPIADS_SIGNIFICANCE

    contests_bonus = 0.0
    contests = student["optional"].get("contests_count")
    if contests is not None and cohort_maxima["contests_count"] > 0:
        contests_bonus = contests / cohort_maxima["contests_count"] * CONTESTS_SIGNIFICANCE

    latest_semester = max(semester_averages)
    latest_semester_sum = sum(grades_by_semester[latest_semester])
    ordered_grades = sorted(
        student["grades"],
        key=lambda item: (
            item["semester"],
            subject_order.get(item["normalized_subject"], 0),
            item["subject"],
        ),
    )
    profile_signature_raw = 0.0
    for index, item in enumerate(ordered_grades, start=1):
        profile_signature_raw += item["grade"] * (
            item["semester"] * 10 + subject_order.get(item["normalized_subject"], 0) + index / 100.0
        )

    # Малый добавочный компонент уменьшает количество совпадений рейтинга,
    # но не должен перекрывать основной вклад оценок и дополнительных факторов.
    collision_reduction_bonus = (
        latest_semester_sum / 10_000.0
        + profile_signature_raw / 1_000_000.0
        + calculate_id_micro_bonus(student["id"])
    )

    base_component = 20.0 * (
        SUBJECT_GRADES_SHARE * weighted_subject_average
        + SEMESTER_AVERAGE_SHARE * weighted_semester_average
        + OVERALL_AVERAGE_SHARE * overall_average
    )
    student["rating"] = round(
        base_component
        + progress_bonus
        + attendance_bonus
        + events_bonus
        + olympiads_bonus
        + contests_bonus
        + collision_reduction_bonus,
        6,
    )
    student["latest_semester_average"] = semester_averages.get(latest_semester, 0.0)


def parse_students_from_workbook(workbook: WorkbookData) -> Tuple[List[dict], List[str]]:
    sheet, header_row = find_input_sheet(workbook)
    headers = read_header_map(sheet, header_row)
    semester_marks = read_semester_marks(sheet, max(1, header_row - 1), sheet["max_col"])
    id_col = next((col for col, header in headers.items() if normalize_text(header) == "id"), None)
    if id_col is None:
        raise ValueError("Колонка id не найдена.")
    track_columns = extract_track_columns(headers)
    track_names = [name for name, _ in track_columns]
    grade_columns = extract_grade_columns(headers, semester_marks)
    optional_columns = detect_optional_columns(headers)
    olympiad_columns = detect_olympiad_columns(headers)

    students: List[dict] = []
    for row in range(header_row + 1, sheet["max_row"] + 1):
        raw_id = get_cell(sheet, row, id_col).strip()
        if not raw_id:
            continue
        grades: List[dict] = []
        for meta in grade_columns:
            raw_grade = get_cell(sheet, row, meta["col"]).strip()
            grade = parse_float(raw_grade)
            if grade is None:
                continue
            grades.append(
                {
                    "semester": meta["semester"],
                    "subject": meta["header"],
                    "normalized_subject": meta["normalized_subject"],
                    "grade": grade,
                }
            )
        if not grades:
            continue
        track_values = {track_name: parse_int(get_cell(sheet, row, col).strip()) for track_name, col in track_columns}
        optional = {key: parse_float(get_cell(sheet, row, col).strip()) for key, col in optional_columns.items()}
        if "no_choice_status" in optional_columns:
            optional["no_choice_status"] = get_cell(sheet, row, optional_columns["no_choice_status"]).strip()
        olympiad_score = calculate_olympiad_score(sheet, row, olympiad_columns) if olympiad_columns else (optional.get("olympiads_count") or 0.0)
        has_track_choice = any(value > 0 for value in normalize_track_preferences(track_values, track_names).values())
        students.append(
            {
                "id": raw_id,
                "grades": grades,
                "track_preferences": normalize_track_preferences(track_values, track_names),
                "optional": optional,
                "olympiad_score": olympiad_score,
                "has_track_choice": has_track_choice,
            }
        )
    if not students:
        raise ValueError("Не найдено ни одной строки со студентами.")
    return students, track_names


def rank_students(students: List[dict]) -> List[dict]:
    semesters = sorted({item["semester"] for student in students for item in student["grades"]})
    weights = compute_semester_weights(semesters)
    subject_order = build_subject_order(students)
    maxima = {
        "events_count": max((student["optional"].get("events_count") or 0.0) for student in students),
        "olympiad_score": max((student.get("olympiad_score") or 0.0) for student in students),
        "contests_count": max((student["optional"].get("contests_count") or 0.0) for student in students),
    }
    for student in students:
        calculate_student_rating(student, weights, maxima, subject_order)
    ranked = sorted(students, key=lambda student: (-student["rating"], -student["latest_semester_average"], student["id"]))
    for index, student in enumerate(ranked, start=1):
        student["rating_place"] = index
    return ranked


def assign_tracks(ranked_students: List[dict], track_names: List[str]) -> List[dict]:
    capacities = build_balanced_capacities(track_names, len(ranked_students))
    loads = {track_name: 0 for track_name in track_names}
    results: List[dict] = []
    for student in ranked_students:
        assigned = choose_track(student["track_preferences"], track_names, loads, capacities)
        loads[assigned] += 1
        results.append(
            {
                "id": student["id"],
                "track": assigned,
                "place": student["rating_place"],
                "rating": student["rating"],
                "has_track_choice": student.get("has_track_choice", True),
                "no_choice_status": student["optional"].get("no_choice_status", ""),
                "track_preferences": student["track_preferences"],
            }
        )
    return results


def xml_cell(ref: str, value: object) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if isinstance(value, float) and (math.isinf(value) or math.isnan(value)):
            value = 0.0
        return f'<c r="{ref}"><v>{value}</v></c>'
    return f'<c r="{ref}" t="inlineStr"><is><t>{escape(str(value))}</t></is></c>'


def build_sheet_xml(rows: Sequence[Sequence[object]]) -> str:
    xml_rows: List[str] = []
    for row_index, row in enumerate(rows, start=1):
        cells = []
        for col_index, value in enumerate(row, start=1):
            cell = xml_cell(f"{index_to_column_letters(col_index)}{row_index}", value)
            if cell:
                cells.append(cell)
        xml_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(xml_rows)}</sheetData></worksheet>'
    )


def write_xlsx_workbook(path: Path, sheets: Sequence[Tuple[str, Sequence[Sequence[object]]]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    created = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    sheet_defs = [(name[:31], rows) for name, rows in sheets]
    workbook_sheet_xml = "".join(
        f'<sheet name="{escape(name)}" sheetId="{index}" r:id="rId{index}"/>'
        for index, (name, _) in enumerate(sheet_defs, start=1)
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheets>{workbook_sheet_xml}</sheets></workbook>"
    )
    sheet_rel_xml = "".join(
        f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.xml"/>'
        for index, _ in enumerate(sheet_defs, start=1)
    )
    workbook_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'{sheet_rel_xml}'
        f'<Relationship Id="rId{len(sheet_defs) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        "</Relationships>"
    )
    sheet_content_types = "".join(
        f'<Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for index, _ in enumerate(sheet_defs, start=1)
    )
    root_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        '</Relationships>'
    )
    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        f"{sheet_content_types}"
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        '</Types>'
    )
    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>'
    )
    app_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Python</Application></Properties>'
    )
    core_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:creator>Codex</dc:creator><cp:lastModifiedBy>Codex</cp:lastModifiedBy>'
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{created}</dcterms:modified></cp:coreProperties>'
    )
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types_xml)
        archive.writestr("_rels/.rels", root_rels_xml)
        archive.writestr("docProps/app.xml", app_xml)
        archive.writestr("docProps/core.xml", core_xml)
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        archive.writestr("xl/styles.xml", styles_xml)
        for index, (_, rows) in enumerate(sheet_defs, start=1):
            archive.writestr(f"xl/worksheets/sheet{index}.xml", build_sheet_xml(rows))


def write_simple_xlsx(path: Path, sheet_name: str, rows: Sequence[Sequence[object]]) -> None:
    write_xlsx_workbook(path, [(sheet_name, rows)])


def create_sample_input_file(path: Path) -> None:
    rows = [
        ["Пример входных данных для распределения по трекам"],
        ["", "1 семестр", "", "", "2 семестр", "", "", "3 семестр", "", "", "", "", "", "", "", "", "", "", "", ""],
        ["id", "Математика", "Программирование", "Английский язык", "Математика", "Алгоритмы", "Проект", "Математика", "Алгоритмы", "Проект", "Трек А", "Трек В", "Трек С", "Посещаемость %", "Внеучебные мероприятия", "Олимпиада 1", "Вес олимпиады 1", "Олимпиада 2", "Вес олимпиады 2", "Конкурсы", "Статус без выбора"],
        ["1001", 4, 4, 5, 4, 5, 4, 5, 5, 5, 1, 2, 3, 96, 3, 1, 1.5, 0, 0, 0, ""],
        ["1002", 5, 5, 5, 5, 4, 5, 5, 5, 5, 0, 0, 0, 90, 1, 1, 2.0, 1, 0.5, 1, "сделает выбор позднее"],
        ["1003", 3, 4, 4, 4, 4, 4, 5, 5, 4, 3, 1, 2, 88, 2, 0, 0, 1, 1.0, 0, ""],
        ["1004", 4, 3, 4, 4, 4, 5, 4, 5, 5, 0, 0, 0, 100, 4, 1, 3.0, 1, 2.0, 1, "уйдет"],
    ]
    write_simple_xlsx(path, "Входные данные", rows)


def write_output_file(path: Path, assignments: List[dict], track_names: List[str]) -> None:
    result_rows: List[List[object]] = [["id ученика", "трек", "место в рейтинге", "рейтинг", "выбор трека", "статус без выбора", "приоритеты"]]
    no_choice_rows: List[List[object]] = [["id ученика", "трек", "место в рейтинге", "рейтинг", "статус без выбора"]]
    assigned_count = {track_name: 0 for track_name in track_names}
    priority_summary_rows: List[List[object]] = [["трек", "назначено студентов", "приоритет 1", "приоритет 1 или 2"]]
    priority_list_rows: List[List[object]] = [["трек", "количество студентов с приоритетом 1 или 2", "id студентов с приоритетом 1 или 2"]]

    for item in assignments:
        assigned_count[item["track"]] += 1
        priorities_text = ", ".join(f"{track_name}: {item['track_preferences'].get(track_name, 0)}" for track_name in track_names)
        choice_mark = "выбор указан" if item["has_track_choice"] else "выбор не указан"
        no_choice_status = item["no_choice_status"] if item["has_track_choice"] else (item["no_choice_status"] or "статус не указан")
        result_rows.append([item["id"], item["track"], item["place"], item["rating"], choice_mark, no_choice_status if not item["has_track_choice"] else "", priorities_text])
        if not item["has_track_choice"]:
            no_choice_rows.append([item["id"], item["track"], item["place"], item["rating"], no_choice_status])

    if len(no_choice_rows) == 1:
        no_choice_rows.append(["нет студентов без выбора", "", "", "", ""])

    for track_name in track_names:
        priority_1_ids = [item["id"] for item in assignments if item["track_preferences"].get(track_name) == 1]
        priority_1_2_ids = [item["id"] for item in assignments if item["track_preferences"].get(track_name) in (1, 2)]
        priority_summary_rows.append([track_name, assigned_count[track_name], len(priority_1_ids), len(priority_1_2_ids)])
        priority_list_rows.append([track_name, len(priority_1_2_ids), ", ".join(priority_1_2_ids)])

    write_xlsx_workbook(
        path,
        [
            ("Результат", result_rows),
            ("Количество_по_трекам", priority_summary_rows),
            ("Без_выбора", no_choice_rows),
            ("Приоритеты_1_2", priority_list_rows),
        ],
    )


def run(input_file: Path, output_folder: Path) -> Tuple[Path, int]:
    workbook = read_xlsx(input_file)
    students, track_names = parse_students_from_workbook(workbook)
    ranked_students = rank_students(students)
    assignments = assign_tracks(ranked_students, track_names)
    output_file = output_folder / OUTPUT_FILE_NAME
    write_output_file(output_file, assignments, track_names)
    return output_file, len(assignments)


def main(input_file_path: str) -> None:
    script_folder = Path(__file__).resolve().parent
    output_folder = script_folder / OUTPUT_FOLDER_NAME
    sample_input_path = script_folder / SAMPLE_INPUT_FILE_NAME
    create_sample_input_file(sample_input_path)
    input_file = Path(input_file_path)
    if not input_file.exists():
        raise SystemExit(
            "Входной файл не найден.\n"
            "Исправьте путь прямо в вызове main(...).\n"
            f"Для примера уже создан файл: {sample_input_path}"
        )
    output_file, count = run(input_file, output_folder)
    print("Готово.")
    print(f"Обработано студентов: {count}")
    print(f"Файл результата: {output_file}")
    print(f"Файл-образец: {sample_input_path}")


if __name__ == "__main__":
    main(r"C:\Users\A006\Downloads\Рейтинг.xlsx")
