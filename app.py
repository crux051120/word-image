"""워드 이미지 자동 삽입 GUI. PyInstaller 로 start.exe 로 빌드한다."""

import os
import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from insert_images import (
    IMAGE_EXTENSIONS,
    STATUS_DUPLICATE_NAME,
    STATUS_NO_IMAGE,
    STATUS_NO_TABLE,
    STATUS_NORMAL,
    collect_images_from_folder,
    create_document_from_images,
    default_created_document_path,
    default_output_path,
    list_document_images,
    process_document,
    validate_table_size,
)

# onefile exe 에서 __file__ 은 임시 압축 해제 폴더라서 exe 위치를 기준으로 삼는다.
BASE_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
DEFAULT_IMAGE_DIR = BASE_DIR / "demo" / "image"

FONT = ("맑은 고딕", 10)
FONT_BOLD = ("맑은 고딕", 10, "bold")
PAD_OUTER = 12
PAD_ROW = 8
COLOR_WARNING = "#C55A00"
COLOR_ERROR = "#C00000"
POLL_INTERVAL_MS = 100

MODE_FOLDER = "folder"
MODE_FILES = "files"
FILTER_ALL = "all"
FILTER_EXCEPTIONS = "exceptions"


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.result_queue: queue.Queue = queue.Queue()
        self.selected_files: list[Path] = []
        self.last_output: Path | None = None
        self.last_inputs: tuple[Path, list[Path], Path] | None = None
        self.last_report = None

        root.title("워드 이미지 자동 삽입")
        root.geometry("760x640")
        root.minsize(720, 600)
        root.option_add("*Font", FONT)
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure(".", font=FONT)
        style.configure("Treeview.Heading", font=FONT_BOLD)

        self.image_mode = tk.StringVar(value=MODE_FOLDER)
        self.image_path = tk.StringVar(value=str(DEFAULT_IMAGE_DIR))
        self.docx_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status_text = tk.StringVar(value="이미지와 Word 문서를 지정한 뒤 실행하세요.")
        self.filter_mode = tk.StringVar(value=FILTER_ALL)

        self.build_layout()

    def build_layout(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True)
        self.insert_tab = ttk.Frame(self.notebook)
        self.create_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.insert_tab, text="기존 문서에 삽입")
        self.notebook.add(self.create_tab, text="이미지로 문서 생성")
        self.creator = CreateDocumentTab(self, self.create_tab)
        frame = ttk.Frame(self.insert_tab, padding=PAD_OUTER)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(6, weight=1)

        # 이미지
        ttk.Label(frame, text="이미지").grid(row=0, column=0, sticky="w", pady=(0, PAD_ROW))
        image_row = ttk.Frame(frame)
        image_row.grid(row=0, column=1, sticky="ew", pady=(0, PAD_ROW))
        image_row.columnconfigure(2, weight=1)
        ttk.Radiobutton(image_row, text="폴더", value=MODE_FOLDER, variable=self.image_mode,
                        command=self.on_mode_change).grid(row=0, column=0, padx=(PAD_ROW, 0))
        ttk.Radiobutton(image_row, text="파일", value=MODE_FILES, variable=self.image_mode,
                        command=self.on_mode_change).grid(row=0, column=1, padx=PAD_ROW)
        self.image_entry = ttk.Entry(image_row, textvariable=self.image_path)
        self.image_entry.grid(row=0, column=2, sticky="ew")
        ttk.Button(frame, text="찾기", command=self.browse_images).grid(
            row=0, column=2, padx=(PAD_ROW, 0), pady=(0, PAD_ROW))

        # Word 문서
        ttk.Label(frame, text="Word 문서").grid(row=1, column=0, sticky="w", pady=(0, PAD_ROW))
        ttk.Entry(frame, textvariable=self.docx_path).grid(
            row=1, column=1, sticky="ew", padx=(PAD_ROW, 0), pady=(0, PAD_ROW))
        ttk.Button(frame, text="찾기", command=self.browse_docx).grid(
            row=1, column=2, padx=(PAD_ROW, 0), pady=(0, PAD_ROW))

        # 결과 파일
        ttk.Label(frame, text="결과 파일").grid(row=2, column=0, sticky="w", pady=(0, PAD_ROW))
        ttk.Entry(frame, textvariable=self.output_path).grid(
            row=2, column=1, sticky="ew", padx=(PAD_ROW, 0), pady=(0, PAD_ROW))
        ttk.Button(frame, text="찾기", command=self.browse_output).grid(
            row=2, column=2, padx=(PAD_ROW, 0), pady=(0, PAD_ROW))

        # 실행
        run_row = ttk.Frame(frame)
        run_row.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(PAD_ROW, PAD_OUTER))
        run_row.columnconfigure(0, weight=1)
        self.status_label = ttk.Label(run_row, textvariable=self.status_text, font=FONT_BOLD)
        self.status_label.grid(row=0, column=0, sticky="w")
        self.run_button = ttk.Button(run_row, text="실행", command=self.run)
        self.run_button.grid(row=0, column=1, sticky="e")

        # 실행 결과 요약
        summary = ttk.LabelFrame(frame, text="실행 결과 요약", padding=PAD_ROW)
        summary.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(0, PAD_OUTER))
        count_row = ttk.Frame(summary)
        count_row.grid(row=0, column=0, sticky="w")
        self.count_labels = {}
        for column, (key, title) in enumerate(
                (("image", "이미지"), ("target", "대상 표"), ("normal", "정상"), ("exception", "예외"))):
            label = ttk.Label(count_row, text=f"{title} -", font=FONT_BOLD)
            label.grid(row=0, column=column, padx=(0, PAD_OUTER * 2))
            self.count_labels[key] = (label, title)
        self.breakdown_label = ttk.Label(summary, text="표만 있음 - · 이미지만 있음 - · 표 이름 중복 -")
        self.breakdown_label.grid(row=1, column=0, sticky="w", pady=(PAD_ROW // 2, 0))
        self.save_label = ttk.Label(summary, text="")
        self.save_label.grid(row=2, column=0, sticky="w", pady=(PAD_ROW // 2, 0))

        # 실행 결과
        result_header = ttk.Frame(frame)
        result_header.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(0, PAD_ROW))
        ttk.Label(result_header, text="실행 결과").grid(row=0, column=0, padx=(0, PAD_OUTER))
        self.filter_all_radio = ttk.Radiobutton(result_header, text="전체", value=FILTER_ALL,
                                                variable=self.filter_mode, command=self.render_rows)
        self.filter_all_radio.grid(row=0, column=1, padx=(0, PAD_ROW))
        self.filter_exception_radio = ttk.Radiobutton(result_header, text="예외만", value=FILTER_EXCEPTIONS,
                                                      variable=self.filter_mode, command=self.render_rows)
        self.filter_exception_radio.grid(row=0, column=2)

        tree_frame = ttk.Frame(frame)
        tree_frame.grid(row=6, column=0, columnspan=3, sticky="nsew")
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        columns = ("table", "name", "image", "status")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        for column, heading, width, stretch in (
            ("table", "표 번호", 70, False),
            ("name", "첫 행 텍스트", 160, True),
            ("image", "매칭 파일", 200, True),
            ("status", "상태", 220, True),
        ):
            self.tree.heading(column, text=heading)
            self.tree.column(column, width=width, stretch=stretch, anchor="w")
        self.tree.tag_configure("exception", foreground=COLOR_WARNING)
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # 하단 버튼
        bottom_row = ttk.Frame(frame)
        bottom_row.grid(row=7, column=0, columnspan=3, sticky="e", pady=(PAD_ROW, 0))
        self.ignore_save_button = ttk.Button(bottom_row, text="예외 무시하고 저장", state="disabled",
                                             command=self.save_ignoring_exceptions)
        self.ignore_save_button.grid(row=0, column=0, padx=(0, PAD_ROW))
        self.open_file_button = ttk.Button(bottom_row, text="결과 파일 열기", state="disabled",
                                           command=lambda: os.startfile(self.last_output))
        self.open_file_button.grid(row=0, column=1, padx=(0, PAD_ROW))
        self.open_folder_button = ttk.Button(bottom_row, text="폴더 열기", state="disabled",
                                             command=lambda: os.startfile(self.last_output.parent))
        self.open_folder_button.grid(row=0, column=2)

    def on_mode_change(self):
        self.selected_files = []
        if self.image_mode.get() == MODE_FOLDER:
            self.image_entry.configure(state="normal")
            self.image_path.set(str(DEFAULT_IMAGE_DIR))
        else:
            # 파일 모드는 여러 파일 목록이라 직접 입력 대신 [찾기]로만 고른다.
            self.image_path.set("[찾기]로 이미지 파일을 선택하세요")
            self.image_entry.configure(state="readonly")

    def browse_images(self):
        if self.image_mode.get() == MODE_FOLDER:
            folder = filedialog.askdirectory(initialdir=self.image_path.get() or str(BASE_DIR))
            if folder:
                self.image_path.set(str(Path(folder)))
            return
        patterns = " ".join(f"*{extension}" for extension in IMAGE_EXTENSIONS)
        files = filedialog.askopenfilenames(initialdir=str(DEFAULT_IMAGE_DIR),
                                            filetypes=[("이미지", patterns)])
        if files:
            self.selected_files = [Path(file) for file in files]
            names = ", ".join(path.name for path in self.selected_files)
            self.image_path.set(f"{len(self.selected_files)}개 파일: {names}")

    def browse_docx(self):
        file = filedialog.askopenfilename(initialdir=str(BASE_DIR), filetypes=[("Word 문서", "*.docx")])
        if file:
            self.docx_path.set(str(Path(file)))
            self.output_path.set(str(default_output_path(Path(file))))

    def browse_output(self):
        initial = Path(self.output_path.get()) if self.output_path.get() else BASE_DIR / "result.docx"
        file = filedialog.asksaveasfilename(initialdir=str(initial.parent), initialfile=initial.name,
                                            defaultextension=".docx", filetypes=[("Word 문서", "*.docx")],
                                            confirmoverwrite=False)
        if file:
            self.output_path.set(str(Path(file)))

    def collect_images(self) -> list[Path] | None:
        if self.image_mode.get() == MODE_FILES:
            if not self.selected_files:
                messagebox.showwarning("입력 확인", "이미지 파일을 선택하세요.")
                return None
            return self.selected_files
        folder = Path(self.image_path.get().strip())
        if not folder.is_dir():
            messagebox.showwarning("입력 확인", f"이미지 폴더가 없습니다.\n{folder}")
            return None
        return collect_images_from_folder(folder)

    def run(self):
        images = self.collect_images()
        if images is None:
            return
        docx_path = Path(self.docx_path.get().strip())
        if not docx_path.is_file():
            messagebox.showwarning("입력 확인", "Word 문서를 지정하세요.")
            return
        if not self.output_path.get().strip():
            messagebox.showwarning("입력 확인", "결과 파일 경로를 지정하세요.")
            return
        output_path = Path(self.output_path.get().strip())
        if output_path.exists() and not messagebox.askyesno(
                "덮어쓰기 확인", f"결과 파일이 이미 있습니다. 덮어쓸까요?\n{output_path}"):
            return

        self.last_inputs = (docx_path, images, output_path)
        self.start_processing(allow_exceptions=False)

    def save_ignoring_exceptions(self):
        # 덮어쓰기 확인은 실행 때 받았으므로 다시 묻지 않는다. 입력을 다시 읽어 화면과 파일을 맞춘다.
        self.start_processing(allow_exceptions=True)

    def start_processing(self, allow_exceptions: bool):
        self.set_running(True)
        threading.Thread(target=self.worker, args=(*self.last_inputs, allow_exceptions), daemon=True).start()
        self.root.after(POLL_INTERVAL_MS, self.poll_result)

    def worker(self, docx_path: Path, images: list[Path], output_path: Path, allow_exceptions: bool):
        try:
            self.result_queue.put(process_document(docx_path, images, output_path, allow_exceptions))
        except Exception as error:  # GUI 스레드에서 오류 메시지로 보여주기 위해 넘긴다.
            self.result_queue.put(error)

    def poll_result(self):
        try:
            result = self.result_queue.get_nowait()
        except queue.Empty:
            self.root.after(POLL_INTERVAL_MS, self.poll_result)
            return
        self.set_running(False)
        if isinstance(result, Exception):
            self.show_status("실패", COLOR_ERROR)
            self.save_label.configure(text=f"오류: {result}", foreground=COLOR_ERROR)
            messagebox.showerror("오류", str(result))
            return
        self.show_report(result)

    def show_report(self, report):
        self.last_report = report
        self.show_summary(report)
        self.render_rows()

        if report.saved:
            self.last_output = report.output_path
            self.open_file_button.configure(state="normal")
            self.open_folder_button.configure(state="normal")
        self.ignore_save_button.configure(state="normal" if report.can_save_ignoring_exceptions else "disabled")

        if report.ok:
            self.show_status("완료", "")
            self.save_label.configure(text=f"저장: {report.output_path.name}", foreground="")
        elif report.saved:
            self.show_status("완료 (예외 무시)", COLOR_WARNING)
            self.save_label.configure(
                text=f"저장됨 (예외 {report.exception_count}건 무시) → {report.output_path.name}",
                foreground=COLOR_WARNING)
        else:
            self.show_status("실패", COLOR_ERROR)
            if report.target_count == 0:
                reason = "2행 1열 표가 없습니다"
            elif report.normal_count == 0:
                reason = "삽입할 수 있는 이미지가 없습니다"
            else:
                reason = f"예외 {report.exception_count}건 — 저장하지 않음"
            self.save_label.configure(text=f"실패: {reason}", foreground=COLOR_ERROR)

    def show_summary(self, report):
        values = {
            "image": report.image_count,
            "target": report.target_count,
            "normal": report.normal_count,
            "exception": report.exception_count,
        }
        for key, (label, title) in self.count_labels.items():
            color = COLOR_WARNING if key == "exception" and values[key] else ""
            label.configure(text=f"{title} {values[key]}", foreground=color)
        self.breakdown_label.configure(
            text=f"표만 있음 {report.count(STATUS_NO_IMAGE)} · "
                 f"이미지만 있음 {report.count(STATUS_NO_TABLE)} · "
                 f"표 이름 중복 {report.count(STATUS_DUPLICATE_NAME)}",
            foreground=COLOR_WARNING if report.exception_count else "")
        self.filter_all_radio.configure(text=f"전체 ({len(report.rows)})")
        self.filter_exception_radio.configure(text=f"예외만 ({report.exception_count})")

    def render_rows(self):
        self.tree.delete(*self.tree.get_children())
        if self.last_report is None:
            return
        only_exceptions = self.filter_mode.get() == FILTER_EXCEPTIONS
        for row in self.last_report.rows:
            if only_exceptions and not row.is_exception:
                continue
            if row.status == STATUS_NORMAL:
                status_text = f"정상 ({row.message})" if row.message else "정상"
            else:
                status_text = f"{row.status} ({row.message})" if row.message else row.status
            table_text = f"표 {row.table_number}" if row.table_number is not None else "-"
            name_text = row.name if row.table_number is not None else "-"
            image_name = row.image_path.name if row.image_path else "(없음)"
            self.tree.insert("", "end", values=(table_text, name_text, image_name, status_text),
                             tags=("exception",) if row.is_exception else ())

    def set_running(self, running: bool):
        self.notebook.tab(self.create_tab, state="disabled" if running else "normal")
        self.run_button.configure(state="disabled" if running else "normal")
        if running:
            self.ignore_save_button.configure(state="disabled")
            self.open_file_button.configure(state="disabled")
            self.open_folder_button.configure(state="disabled")
            self.show_status("처리 중...", "")

    def show_status(self, text: str, color: str):
        self.status_text.set(text)
        self.status_label.configure(foreground=color)


class CreateDocumentTab:
    """새 문서 탭의 입력·표시만 담당하며 문서 처리는 insert_images에 위임한다."""

    def __init__(self, app: App, parent):
        self.app = app
        self.root = app.root
        self.folder = tk.StringVar(value=str(DEFAULT_IMAGE_DIR))
        self.suggested_output = str(default_created_document_path(DEFAULT_IMAGE_DIR))
        self.output = tk.StringVar(value=self.suggested_output)
        self.width = tk.StringVar(value="16")
        self.height = tk.StringVar(value="10")
        self.summary = tk.StringVar(value="이미지 -  ·  생성 표 -  ·  정상 -  ·  오류 -")
        self.status = tk.StringVar(value="이미지 폴더를 지정하고 목록을 확인하세요.")
        self.images = []
        self.loaded_folder = None
        self.last_inputs = None
        self.last_output = None
        self.result_queue = queue.Queue()
        self.input_widgets = []

        frame = ttk.Frame(parent, padding=PAD_OUTER)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(6, weight=1)
        for row, label, variable, command in (
                (0, "이미지 폴더", self.folder, self.browse_folder),
                (1, "결과 파일", self.output, self.browse_output)):
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=(0, PAD_ROW))
            entry = ttk.Entry(frame, textvariable=variable)
            entry.grid(row=row, column=1, sticky="ew", padx=PAD_ROW, pady=(0, PAD_ROW))
            button = ttk.Button(frame, text="찾기", command=command)
            button.grid(row=row, column=2, pady=(0, PAD_ROW))
            self.input_widgets.extend((entry, button))

        size_row = ttk.Frame(frame)
        size_row.grid(row=2, column=0, columnspan=3, sticky="w", pady=(0, PAD_ROW))
        for column, (label, variable) in enumerate((("표 너비 (cm)", self.width), ("전체 높이 (cm)", self.height))):
            ttk.Label(size_row, text=label).grid(row=0, column=column * 2, padx=(0, PAD_ROW))
            entry = ttk.Entry(size_row, textvariable=variable, width=7)
            entry.grid(row=0, column=column * 2 + 1, padx=(0, PAD_OUTER))
            self.input_widgets.append(entry)
        ttk.Label(frame, text="A4 세로 · 여백 2cm · 이름 행 1cm 포함 · 이미지 비율 유지\n"
                              "너비 1~17cm / 높이 1.5~25.5cm · PNG/JPG/JPEG/GIF/BMP · 하위 폴더 제외").grid(
            row=3, column=0, columnspan=3, sticky="w", pady=(0, PAD_ROW))

        actions = ttk.Frame(frame)
        actions.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(0, PAD_OUTER))
        actions.columnconfigure(0, weight=1)
        self.status_label = ttk.Label(actions, textvariable=self.status, font=FONT_BOLD, wraplength=450)
        self.status_label.grid(row=0, column=0, sticky="w")
        self.refresh_button = ttk.Button(actions, text="목록 확인", command=self.refresh)
        self.refresh_button.grid(row=0, column=1, padx=(PAD_ROW, PAD_ROW))
        self.run_button = ttk.Button(actions, text="문서 생성", command=self.run)
        self.run_button.grid(row=0, column=2)
        self.input_widgets.extend((self.refresh_button, self.run_button))
        ttk.Label(frame, textvariable=self.summary, font=FONT_BOLD).grid(
            row=5, column=0, columnspan=3, sticky="w", pady=(0, PAD_ROW))

        tree_frame = ttk.Frame(frame)
        tree_frame.grid(row=6, column=0, columnspan=3, sticky="nsew")
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(tree_frame, columns=("number", "file", "status"), show="headings")
        for column, heading, width, stretch in (
                ("number", "번호", 50, False), ("file", "이미지 파일", 260, True), ("status", "처리 상태", 330, True)):
            self.tree.heading(column, text=heading)
            self.tree.column(column, width=width, stretch=stretch, anchor="w")
        self.tree.tag_configure("error", foreground=COLOR_ERROR)
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        horizontal = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=horizontal.set)
        horizontal.grid(row=1, column=0, sticky="ew")
        bottom = ttk.Frame(frame)
        bottom.grid(row=7, column=0, columnspan=3, sticky="e", pady=(PAD_ROW, 0))
        self.ignore_button = ttk.Button(bottom, text="오류 제외하고 저장", state="disabled",
                                        command=lambda: self.run(allow_exceptions=True))
        self.open_button = ttk.Button(bottom, text="결과 파일 열기", state="disabled",
                                      command=lambda: os.startfile(self.last_output))
        self.folder_button = ttk.Button(bottom, text="폴더 열기", state="disabled",
                                        command=lambda: os.startfile(self.last_output.parent))
        for column, button in enumerate((self.ignore_button, self.open_button, self.folder_button)):
            button.grid(row=0, column=column, padx=(PAD_ROW, 0))
        for variable in (self.folder, self.output, self.width, self.height):
            variable.trace_add("write", self.input_changed)

    def input_changed(self, *_):
        # 입력이 바뀐 뒤 이전 결과를 새 입력의 결과로 오인해 저장하지 않게 한다.
        self.reset_result()
        self.status.set("입력이 변경되었습니다. 목록 확인 또는 문서 생성을 실행하세요.")
        self.status_label.configure(foreground="")
        self.summary.set(f"이미지 {len(self.images)}  ·  생성 표 -  ·  정상 -  ·  오류 -")
        self.tree.delete(*self.tree.get_children())

    def browse_folder(self):
        folder = filedialog.askdirectory(initialdir=self.folder.get() or str(BASE_DIR))
        if folder:
            self.folder.set(folder)
            self.refresh()

    def browse_output(self):
        initial = Path(self.output.get()) if self.output.get() else BASE_DIR / "images.docx"
        result = filedialog.asksaveasfilename(initialdir=str(initial.parent), initialfile=initial.name,
                                             defaultextension=".docx", filetypes=[("Word 문서", "*.docx")],
                                             confirmoverwrite=False)
        if result:
            self.output.set(result)

    def reset_result(self):
        self.last_output = None
        self.last_inputs = None
        for button in (self.ignore_button, self.open_button, self.folder_button):
            button.configure(state="disabled")

    def refresh(self):
        self.reset_result()
        self.images = []
        self.tree.delete(*self.tree.get_children())
        self.summary.set("이미지 -  ·  생성 표 -  ·  정상 -  ·  오류 -")
        folder = Path(self.folder.get().strip())
        try:
            if not self.folder.get().strip() or not folder.is_dir():
                raise ValueError("이미지 폴더를 지정하세요.")
            self.images = list_document_images(folder)
        except (OSError, ValueError) as error:
            self.status.set(f"오류: {error}")
            self.status_label.configure(foreground=COLOR_ERROR)
            return False
        if folder.resolve() != self.loaded_folder:
            suggested = str(default_created_document_path(folder))
            if not self.output.get().strip() or self.output.get() == self.suggested_output:
                self.output.set(suggested)
            self.suggested_output = suggested
            self.loaded_folder = folder.resolve()
        for number, path in enumerate(self.images, 1):
            self.tree.insert("", "end", values=(number, path.name, "대기"))
        self.summary.set(f"이미지 {len(self.images)}  ·  생성 표 -  ·  정상 -  ·  오류 -")
        if not self.images:
            self.summary.set("이미지 0  ·  생성 표 0  ·  정상 0  ·  오류 0")
        self.status.set("목록 확인 완료" if self.images else "지원하는 이미지가 없습니다. 저장하지 않습니다.")
        self.status_label.configure(foreground="" if self.images else COLOR_ERROR)
        return bool(self.images)

    def run(self, allow_exceptions=False):
        if allow_exceptions:
            inputs = self.last_inputs
            if inputs is None:
                return
        else:
            if not self.refresh():
                return
            try:
                width, height = float(self.width.get()), float(self.height.get())
                validate_table_size(width, height)
                if not self.output.get().strip():
                    raise ValueError("결과 파일 경로를 지정하세요.")
                output = Path(self.output.get().strip())
                if output.suffix.lower() != ".docx":
                    raise ValueError("결과 파일 확장자는 .docx여야 합니다.")
                if not output.parent.is_dir():
                    raise ValueError("결과 파일의 저장 폴더가 없습니다.")
            except ValueError as error:
                self.status.set(f"입력 확인: {error}")
                self.status_label.configure(foreground=COLOR_ERROR)
                messagebox.showwarning("입력 확인", str(error))
                return
            inputs = (list(self.images), output, width, height)
        if inputs[1].exists() and not messagebox.askyesno(
                "덮어쓰기 확인", f"결과 파일이 이미 있습니다. 덮어쓸까요?\n{inputs[1]}"):
            return
        self.last_inputs = inputs
        self.last_output = None
        self.set_running(True)
        threading.Thread(target=self.worker, args=(*inputs, allow_exceptions), daemon=True).start()
        self.root.after(POLL_INTERVAL_MS, self.poll_result)

    def set_running(self, running):
        self.app.notebook.tab(self.app.insert_tab, state="disabled" if running else "normal")
        for widget in self.input_widgets:
            widget.configure(state="disabled" if running else "normal")
        if running:
            for button in (self.ignore_button, self.open_button, self.folder_button):
                button.configure(state="disabled")
            self.status.set("처리 중...")
            self.status_label.configure(foreground="")

    def worker(self, images, output, width, height, allow_exceptions):
        try:
            result = create_document_from_images(images, output, width, height, allow_exceptions)
        except Exception as error:
            result = error
        self.result_queue.put(result)

    def poll_result(self):
        try:
            report = self.result_queue.get_nowait()
        except queue.Empty:
            self.root.after(POLL_INTERVAL_MS, self.poll_result)
            return
        self.set_running(False)
        if isinstance(report, Exception):
            self.status.set(f"실패: {report}")
            self.status_label.configure(foreground=COLOR_ERROR)
            self.summary.set(f"이미지 {len(self.last_inputs[0])}  ·  생성 표 -  ·  정상 -  ·  오류 -")
            return
        self.tree.delete(*self.tree.get_children())
        for number, row in enumerate(report.rows, 1):
            detail = f"{row.status}: {row.message}" if row.message else row.status
            self.tree.insert("", "end", values=(number, row.image_path.name, detail),
                             tags=("error",) if row.is_exception else ())
        self.summary.set(f"이미지 {report.image_count}  ·  생성 표 {report.target_count}  ·  "
                         f"정상 {report.normal_count}  ·  오류 {report.exception_count}")
        self.ignore_button.configure(state="normal" if report.can_save_ignoring_exceptions else "disabled")
        if report.saved:
            self.last_output = report.output_path
            for button in (self.open_button, self.folder_button):
                button.configure(state="normal")
            prefix = f"완료 (오류 {report.exception_count}건 제외)" if report.exception_count else "완료"
            self.status.set(f"{prefix}: {report.output_path.name}")
            self.status_label.configure(foreground=COLOR_WARNING if report.exception_count else "")
        else:
            self.status.set(f"실패: 오류 {report.exception_count}건, 저장하지 않았습니다.")
            self.status_label.configure(foreground=COLOR_ERROR)


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
