"""docx 의 2행 1열 표를 찾아, 첫 행에 적힌 이름의 이미지를 두 번째 행에 삽입한다."""

from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Cm

# 같은 이름에 여러 확장자가 있을 때 이 순서로 고른다.
IMAGE_EXTENSIONS = (".png", ".jpg", ".jpeg", ".gif", ".bmp")

# 셀 좌우 기본 여백(약 0.19cm × 2)만큼 줄여야 이미지가 셀 밖으로 밀려나지 않는다.
CELL_PADDING = Cm(0.4)

STATUS_NORMAL = "정상"
STATUS_NO_IMAGE = "예외: 이미지 없음"
STATUS_NO_TABLE = "예외: 해당 표 없음"
STATUS_DUPLICATE_NAME = "예외: 표 이름 중복"
EXCEPTION_STATUSES = (STATUS_NO_IMAGE, STATUS_NO_TABLE, STATUS_DUPLICATE_NAME)


@dataclass
class ResultRow:
    table_number: int | None
    name: str
    image_path: Path | None
    status: str
    message: str = ""

    @property
    def is_exception(self) -> bool:
        return self.status in EXCEPTION_STATUSES


@dataclass
class InsertReport:
    output_path: Path
    image_count: int = 0
    target_count: int = 0
    rows: list[ResultRow] = field(default_factory=list)
    saved: bool = False

    def count(self, status: str) -> int:
        return sum(1 for row in self.rows if row.status == status)

    @property
    def normal_count(self) -> int:
        return self.count(STATUS_NORMAL)

    @property
    def exception_count(self) -> int:
        return sum(1 for row in self.rows if row.is_exception)

    @property
    def has_content(self) -> bool:
        # 대상 표 0개나 정상 0건이면 저장할 것이 없다.
        return self.target_count > 0 and self.normal_count > 0

    @property
    def ok(self) -> bool:
        return self.has_content and self.exception_count == 0

    @property
    def can_save_ignoring_exceptions(self) -> bool:
        return not self.saved and self.has_content and self.exception_count > 0


def collect_images_from_folder(folder: Path) -> list[Path]:
    return [path for path in Path(folder).iterdir() if path.is_file()]


def build_image_index(image_paths: list[Path]) -> dict[str, list[Path]]:
    """소문자 파일 이름(확장자 제외) → 확장자 우선순위로 정렬된 경로 목록."""
    index: dict[str, list[Path]] = {}
    for path in image_paths:
        path = Path(path)
        extension = path.suffix.lower()
        if extension not in IMAGE_EXTENSIONS:
            continue
        index.setdefault(path.stem.strip().lower(), []).append(path)
    for paths in index.values():
        paths.sort(key=lambda path: IMAGE_EXTENSIONS.index(path.suffix.lower()))
    return index


def is_target_table(table) -> bool:
    grid_columns = table._tbl.tblGrid.findall(qn("w:gridCol"))
    return len(table.rows) == 2 and len(grid_columns) == 1


def available_width(document, table):
    column_width = table.columns[0].width
    if column_width:
        return column_width - CELL_PADDING
    section = document.sections[-1]
    return section.page_width - section.left_margin - section.right_margin - CELL_PADDING


def replace_cell_with_image(cell, image_path: Path, width) -> None:
    # 셀 속성(tcPr)만 남기고 기존 문단·중첩 표를 모두 지운다.
    for child in list(cell._tc):
        if child.tag != qn("w:tcPr"):
            cell._tc.remove(child)
    paragraph = cell.add_paragraph()
    paragraph.add_run().add_picture(str(image_path), width=width)


def process_document(docx_path: Path, image_paths: list[Path], output_path: Path,
                     allow_exceptions: bool = False) -> InsertReport:
    """예외가 하나라도 있으면 저장하지 않는다. allow_exceptions=True 면 정상 표만 채워 저장한다."""
    docx_path = Path(docx_path)
    output_path = Path(output_path)
    if docx_path.resolve() == output_path.resolve():
        raise ValueError("결과 파일이 원본 Word 문서와 같습니다. 다른 경로를 지정하세요.")

    document = Document(str(docx_path))
    image_index = build_image_index(image_paths)
    report = InsertReport(output_path=output_path,
                          image_count=sum(len(paths) for paths in image_index.values()))

    targets = [
        (table_number, table, table.rows[0].cells[0].text.strip())
        for table_number, table in enumerate(document.tables, start=1)
        if is_target_table(table)
    ]
    report.target_count = len(targets)
    numbers_by_name: dict[str, list[int]] = {}
    for table_number, _, name in targets:
        if name:
            numbers_by_name.setdefault(name.lower(), []).append(table_number)

    for table_number, table, name in targets:
        key = name.lower()
        candidates = image_index.get(key, [])
        if name and len(numbers_by_name[key]) > 1:
            tables = ", ".join(str(number) for number in numbers_by_name[key])
            report.rows.append(ResultRow(table_number, name, candidates[0] if candidates else None,
                                         STATUS_DUPLICATE_NAME, f"표 {tables}"))
            continue
        if not name or not candidates:
            message = "첫 행이 비어 있음" if not name else ""
            report.rows.append(ResultRow(table_number, name, None, STATUS_NO_IMAGE, message))
            continue

        image_path = candidates[0]
        replace_cell_with_image(table.rows[1].cells[0], image_path, available_width(document, table))
        others = ", ".join(path.name for path in candidates[1:])
        report.rows.append(ResultRow(table_number, name, image_path, STATUS_NORMAL,
                                     f"중복 무시: {others}" if others else ""))

    unused_images = sorted(
        (path for key, paths in image_index.items() if key not in numbers_by_name for path in paths),
        key=lambda path: path.name.lower(),
    )
    for path in unused_images:
        report.rows.append(ResultRow(None, "", path, STATUS_NO_TABLE))

    if report.has_content and (report.exception_count == 0 or allow_exceptions):
        document.save(str(output_path))
        report.saved = True
    return report


def default_output_path(docx_path: Path) -> Path:
    docx_path = Path(docx_path)
    return docx_path.with_name(f"{docx_path.stem}_filled{docx_path.suffix}")
