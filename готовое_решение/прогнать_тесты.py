from __future__ import annotations

import importlib.util
import shutil
from collections import Counter
from pathlib import Path


ROOT = Path(__file__).resolve().parent
TEST_ROOT = ROOT / "тест"
SOLUTION_PATH = ROOT / "распределение_по_трекам.py"
SOURCE_INPUT = Path(r"C:\Users\A006\Downloads\Рейтинг.xlsx")


def load_solution_module():
    spec = importlib.util.spec_from_file_location("track_solution", SOLUTION_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec and spec.loader
    spec.loader.exec_module(module)
    return module


def prepare_test_root(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)
    for item in path.iterdir():
        if item.is_dir():
            shutil.rmtree(item)
        else:
            item.unlink()


def full_headers() -> list[str]:
    return [
        "id",
        "Математика",
        "Программирование",
        "Английский язык",
        "Математика",
        "Алгоритмы",
        "Проект",
        "Математика",
        "Алгоритмы",
        "Проект",
        "Трек А",
        "Трек В",
        "Трек С",
        "Посещаемость %",
        "Внеучебные мероприятия",
        "Олимпиада 1",
        "Вес олимпиады 1",
        "Олимпиада 2",
        "Вес олимпиады 2",
        "Конкурсы",
        "Статус без выбора",
    ]


def minimal_headers() -> list[str]:
    return [
        "id",
        "Математика",
        "Программирование",
        "Английский язык",
        "Математика",
        "Алгоритмы",
        "Проект",
        "Математика",
        "Алгоритмы",
        "Проект",
        "Трек А",
        "Трек В",
        "Трек С",
    ]


def build_rows(title: str, headers: list[str], students: list[list[object]]) -> list[list[object]]:
    return [
        [title],
        ["", "1 семестр", "", "", "2 семестр", "", "", "3 семестр", "", "", "", "", "", "", "", "", "", "", ""],
        headers,
        *students,
    ]


def read_result_rows(module, path: Path) -> list[dict]:
    workbook = module.read_xlsx(path)
    sheet = workbook.sheets[0]
    rows: list[dict] = []
    for row in range(2, sheet["max_row"] + 1):
        student_id = module.get_cell(sheet, row, 1)
        if not student_id:
            continue
        rows.append(
            {
                "id": student_id,
                "track": module.get_cell(sheet, row, 2),
                "place": module.get_cell(sheet, row, 3),
                "rating": module.get_cell(sheet, row, 4),
            }
        )
    return rows


def rating_collision_count(result_rows: list[dict]) -> tuple[int, int]:
    ratings = [str(row["rating"]) for row in result_rows]
    counter = Counter(ratings)
    unique_count = len(counter)
    collisions = sum(count - 1 for count in counter.values() if count > 1)
    return unique_count, collisions


def preview_lines(result_rows: list[dict], limit: int = 4) -> list[str]:
    lines = []
    for row in result_rows[:limit]:
        lines.append(f"- {row['id']}: {row['track']}, место {row['place']}, рейтинг {row['rating']}")
    return lines


def write_case(module, case_dir: Path, description: str, rows: list[list[object]]) -> tuple[list[dict], int]:
    case_dir.mkdir(parents=True, exist_ok=True)
    input_path = case_dir / "вход.xlsx"
    (case_dir / "описание.txt").write_text(description, encoding="utf-8")
    module.write_simple_xlsx(input_path, "Тест", rows)
    output_path, count = module.run(input_path, case_dir / "выход")
    return read_result_rows(module, output_path), count


def main() -> None:
    module = load_solution_module()
    prepare_test_root(TEST_ROOT)

    summary = ["# Сводка по тестам", ""]

    source_case = TEST_ROOT / "01_исходный_файл"
    source_case.mkdir(parents=True, exist_ok=True)
    copied_input = source_case / "Рейтинг.xlsx"
    shutil.copy2(SOURCE_INPUT, copied_input)
    (source_case / "описание.txt").write_text(
        "Проверка на исходном файле без изменения структуры данных.",
        encoding="utf-8",
    )
    output_path, count = module.run(copied_input, source_case / "выход")
    result_rows = read_result_rows(module, output_path)
    unique_count, collisions = rating_collision_count(result_rows)
    summary += [
        "## 01. Исходный файл",
        f"- Обработано студентов: {count}",
        f"- Уникальных значений рейтинга: {unique_count}",
        f"- Совпадений рейтинга: {collisions}",
        *preview_lines(result_rows),
        "",
    ]

    cases = [
        (
            "02_антисовпадение_одинаковых_оценок",
            "Проверка, что при максимально похожих оценках рейтинг все равно имеет минимум совпадений.",
            build_rows(
                "Проверка антисовпадения",
                full_headers(),
                [
                    ["2101", 5, 5, 5, 5, 5, 5, 5, 5, 5, 1, 2, 3, 90, 1, 0, 0, 0, 0, 0],
                    ["2102", 5, 5, 5, 5, 5, 5, 5, 5, 5, 1, 2, 3, 90, 1, 0, 0, 0, 0, 0],
                    ["2103", 5, 5, 5, 5, 5, 5, 5, 5, 5, 1, 2, 3, 90, 1, 0, 0, 0, 0, 0],
                    ["2104", 5, 5, 5, 5, 5, 5, 5, 5, 5, 1, 2, 3, 90, 1, 0, 0, 0, 0, 0],
                ],
            ),
        ),
        (
            "03_нули_в_приоритетах",
            "Проверка, что 0 в приоритетах трека корректно трактуется как 'без разницы'.",
            build_rows(
                "Проверка нулевых приоритетов",
                full_headers(),
                [
                    ["3101", 4, 4, 4, 4, 4, 4, 4, 4, 4, 0, 0, 0, 90, 0, 0, 0, 0, 0, 0],
                    ["3102", 5, 5, 5, 5, 5, 5, 5, 5, 5, 0, 0, 0, 90, 0, 0, 0, 0, 0, 0, "сделает выбор позднее"],
                    ["3103", 3, 3, 3, 3, 3, 3, 3, 3, 3, 0, 0, 0, 90, 0, 0, 0, 0, 0, 0, "уйдет"],
                    ["3104", 4, 5, 4, 4, 5, 4, 4, 5, 4, 0, 0, 0, 90, 0, 0, 0, 0, 0, 0],
                    ["3105", 5, 4, 5, 5, 4, 5, 5, 4, 5, 0, 0, 0, 90, 0, 0, 0, 0, 0, 0],
                    ["3106", 3, 4, 3, 3, 4, 3, 3, 4, 3, 0, 0, 0, 90, 0, 0, 0, 0, 0, 0],
                ],
            ),
        ),
        (
            "04_подтвержденный_прогресс",
            "Проверка, что рост в поздних семестрах повышает рейтинг, а снижение уменьшает его.",
            build_rows(
                "Проверка прогресса",
                full_headers(),
                [
                    ["4101", 3, 3, 4, 4, 4, 4, 5, 5, 5, 1, 2, 3, 90, 0, 0, 0, 0, 0, 0],
                    ["4102", 5, 5, 4, 4, 4, 4, 3, 3, 3, 1, 2, 3, 90, 0, 0, 0, 0, 0, 0],
                    ["4103", 4, 4, 4, 4, 4, 4, 4, 4, 4, 2, 1, 3, 90, 0, 0, 0, 0, 0, 0],
                ],
            ),
        ),
        (
            "05_вес_поздних_семестров",
            "Проверка, что более поздние семестры сильнее влияют на рейтинг.",
            build_rows(
                "Проверка весов семестров",
                full_headers(),
                [
                    ["5101", 5, 5, 5, 4, 4, 4, 3, 3, 3, 1, 2, 3, 90, 0, 0, 0, 0, 0, 0],
                    ["5102", 3, 3, 3, 4, 4, 4, 5, 5, 5, 1, 2, 3, 90, 0, 0, 0, 0, 0, 0],
                    ["5103", 4, 4, 4, 4, 4, 4, 4, 4, 4, 2, 1, 3, 90, 0, 0, 0, 0, 0, 0],
                ],
            ),
        ),
        (
            "06_олимпиады_с_весами",
            "Проверка, что разные веса олимпиад дают разный итоговый вклад в рейтинг.",
            build_rows(
                "Проверка олимпиад с весами",
                full_headers(),
                [
                    ["6101", 5, 5, 5, 5, 5, 5, 5, 5, 5, 1, 2, 3, 90, 0, 1, 3.0, 1, 2.0, 0],
                    ["6102", 5, 5, 5, 5, 5, 5, 5, 5, 5, 1, 2, 3, 90, 0, 1, 1.0, 0, 0, 0],
                    ["6103", 5, 5, 5, 5, 5, 5, 5, 5, 5, 2, 1, 3, 90, 0, 0, 0, 0, 0, 0],
                ],
            ),
        ),
        (
            "07_без_дополнительных_полей",
            "Проверка, что отсутствие дополнительных колонок не мешает расчету и все лишнее корректно игнорируется.",
            build_rows(
                "Проверка без дополнительных полей",
                minimal_headers(),
                [
                    ["7101", 4, 4, 4, 4, 4, 4, 4, 4, 4, 1, 2, 3],
                    ["7102", 5, 5, 5, 5, 5, 5, 5, 5, 5, 2, 1, 3],
                    ["7103", 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 1, 2],
                    ["7104", 4, 5, 4, 4, 5, 4, 4, 5, 4, 0, 0, 0],
                ],
            ),
        ),
    ]

    for case_name, description, rows in cases:
        result_rows, count = write_case(module, TEST_ROOT / case_name, description, rows)
        unique_count, collisions = rating_collision_count(result_rows)
        summary += [
            f"## {case_name.replace('_', ' ')}",
            f"- Обработано студентов: {count}",
            f"- Уникальных значений рейтинга: {unique_count}",
            f"- Совпадений рейтинга: {collisions}",
            *preview_lines(result_rows),
            "",
        ]

    (TEST_ROOT / "сводка_по_тестам.md").write_text("\n".join(summary), encoding="utf-8")
    print(f"Готово. Тесты записаны в: {TEST_ROOT}")


if __name__ == "__main__":
    main()
