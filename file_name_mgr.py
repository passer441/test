from __future__ import annotations

import os
import time
import tkinter as tk
import unicodedata
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Optional

import pyperclip
import pythoncom
import win32com.client


APP_NAME = "OLED 측정 도우미"
DEFAULT_EXTENSIONS = ".csv"
CHECK_INTERVAL_MS = 1000


def normalize_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_filename(name: str) -> str:
    """경로를 제거하고 파일명 전체를 정규화합니다."""
    text = unicodedata.normalize("NFKC", str(name)).strip()
    text = text.replace("\\", "/")
    return Path(text).name.strip().casefold()


def normalize_stem(name: str) -> str:
    """경로와 마지막 확장자를 제거한 파일명을 정규화합니다."""
    return Path(normalize_filename(name)).stem.strip().casefold()


class MeasurementHelper:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("640x560")
        self.root.minsize(620, 520)

        self.excel = None
        self.workbook = None
        self.sheet = None
        self.user_connected = False

        self.previous_clipboard = ""
        self.running = True

        self.folder_var = tk.StringVar()
        self.filename_col_var = tk.IntVar(value=1)
        self.status_col_var = tk.IntVar(value=2)
        self.auto_status_col_var = tk.BooleanVar(value=True)
        self.start_row_var = tk.IntVar(value=2)
        self.extensions_var = tk.StringVar(value=DEFAULT_EXTENSIONS)

        self.connection_text_var = tk.StringVar(value="연결 안 됨")
        self.connection_detail_var = tk.StringVar(
            value="실행 중인 Excel을 찾는 중입니다."
        )

        self.clipboard_var = tk.StringVar(value="-")
        self.current_var = tk.StringVar(value="-")
        self.next_var = tk.StringVar(value="-")
        self.saved_var = tk.StringVar(value="-")
        self.state_var = tk.StringVar(value="대기")
        self.progress_text_var = tk.StringVar(value="0 / 0")
        self.footer_var = tk.StringVar(
            value="엑셀과 측정 폴더를 1초마다 확인합니다."
        )

        self._build_ui()
        self._load_settings()

        self.filename_col_var.trace_add(
            "write",
            self.on_filename_col_changed,
        )
        self.on_auto_status_toggle()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(CHECK_INTERVAL_MS, self.refresh_all)

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill="both", expand=True)

        connection_box = ttk.LabelFrame(main, text="엑셀 연결 상태", padding=10)
        connection_box.pack(fill="x", pady=(0, 8))

        self.connection_indicator = tk.Label(
            connection_box,
            text="●",
            fg="#c62828",
            font=("Arial", 20, "bold"),
        )
        self.connection_indicator.grid(row=0, column=0, rowspan=2, padx=(0, 10))

        self.connection_status_label = tk.Label(
            connection_box,
            textvariable=self.connection_text_var,
            fg="#c62828",
            font=("맑은 고딕", 12, "bold"),
        )
        self.connection_status_label.grid(
            row=0, column=1, sticky="w"
        )

        ttk.Label(
            connection_box,
            textvariable=self.connection_detail_var,
            wraplength=400,
        ).grid(row=1, column=1, sticky="w")

        ttk.Button(
            connection_box,
            text="현재 엑셀에 연결",
            command=self.connect_excel,
        ).grid(row=0, column=2, rowspan=2, padx=(10, 0))

        connection_box.columnconfigure(1, weight=1)

        settings_box = ttk.LabelFrame(main, text="설정", padding=10)
        settings_box.pack(fill="x", pady=(0, 8))

        ttk.Label(settings_box, text="측정 폴더").grid(
            row=0, column=0, sticky="w", pady=3
        )

        folder_frame = ttk.Frame(settings_box)
        folder_frame.grid(row=0, column=1, sticky="ew", pady=3)

        ttk.Entry(
            folder_frame,
            textvariable=self.folder_var,
        ).pack(side="left", fill="x", expand=True)

        ttk.Button(
            folder_frame,
            text="폴더 선택",
            command=self.choose_folder,
        ).pack(side="left", padx=(5, 0))

        ttk.Label(settings_box, text="파일 이름 열 번호").grid(
            row=1, column=0, sticky="w", pady=3
        )
        ttk.Spinbox(
            settings_box,
            from_=1,
            to=16384,
            textvariable=self.filename_col_var,
            width=10,
        ).grid(row=1, column=1, sticky="w", pady=3)

        ttk.Label(settings_box, text="상태 열 번호").grid(
            row=2, column=0, sticky="w", pady=3
        )

        status_col_frame = ttk.Frame(settings_box)
        status_col_frame.grid(row=2, column=1, sticky="w", pady=3)

        self.status_col_spinbox = ttk.Spinbox(
            status_col_frame,
            from_=1,
            to=16384,
            textvariable=self.status_col_var,
            width=10,
        )
        self.status_col_spinbox.pack(side="left")

        ttk.Checkbutton(
            status_col_frame,
            text="자동: 파일 이름 열 + 1",
            variable=self.auto_status_col_var,
            command=self.on_auto_status_toggle,
        ).pack(side="left", padx=(8, 0))

        ttk.Label(settings_box, text="시작 행").grid(
            row=3, column=0, sticky="w", pady=3
        )
        ttk.Spinbox(
            settings_box,
            from_=1,
            to=1048576,
            textvariable=self.start_row_var,
            width=10,
        ).grid(row=3, column=1, sticky="w", pady=3)

        ttk.Label(settings_box, text="사용 확장자").grid(
            row=4, column=0, sticky="w", pady=3
        )
        ttk.Entry(
            settings_box,
            textvariable=self.extensions_var,
        ).grid(row=4, column=1, sticky="ew", pady=3)

        ttk.Label(
            settings_box,
            text=(
                "공백: 엑셀 파일 목록에 확장자를 포함합니다. 예) Sample001.csv\n"
                ".csv 입력: 엑셀 파일 목록에는 확장자를 제외합니다. 예) Sample001"
            ),
            foreground="#555555",
            justify="left",
        ).grid(
            row=5,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(2, 5),
        )

        settings_box.columnconfigure(1, weight=1)

        status_box = ttk.LabelFrame(main, text="측정 상태", padding=10)
        status_box.pack(fill="x", pady=(0, 8))

        rows = [
            ("현재 클립보드", self.clipboard_var),
            ("현재 측정 대상", self.current_var),
            ("다음 샘플", self.next_var),
            ("최근 확인 파일", self.saved_var),
            ("상태", self.state_var),
        ]

        for i, (title, variable) in enumerate(rows):
            ttk.Label(status_box, text=title).grid(
                row=i, column=0, sticky="w", pady=3
            )
            ttk.Label(
                status_box,
                textvariable=variable,
            ).grid(row=i, column=1, sticky="w", pady=3)

        status_box.columnconfigure(1, weight=1)

        progress_box = ttk.LabelFrame(main, text="진행률", padding=10)
        progress_box.pack(fill="x", pady=(0, 8))

        self.progress = ttk.Progressbar(
            progress_box,
            orient="horizontal",
            mode="determinate",
            maximum=100,
        )
        self.progress.pack(fill="x")

        ttk.Label(
            progress_box,
            textvariable=self.progress_text_var,
            anchor="center",
        ).pack(fill="x", pady=(5, 0))

        button_frame = ttk.Frame(main)
        button_frame.pack(fill="x", pady=(0, 8))

        ttk.Button(
            button_frame,
            text="현재 대상 다시 복사",
            command=self.copy_current_target,
        ).pack(side="left", fill="x", expand=True)

        ttk.Button(
            button_frame,
            text="즉시 새로고침",
            command=self.refresh_all,
        ).pack(side="left", fill="x", expand=True, padx=(5, 0))

        ttk.Label(
            main,
            textvariable=self.footer_var,
            wraplength=600,
        ).pack(fill="x")

    def choose_folder(self) -> None:
        start_folder = self.folder_var.get().strip() or str(Path.home())
        folder = filedialog.askdirectory(
            title="측정 폴더 선택",
            initialdir=start_folder,
        )
        if folder:
            self.folder_var.set(folder)
            self._save_settings()
            self.refresh_all()

    def on_filename_col_changed(self, *_args) -> None:
        if not self.auto_status_col_var.get():
            return

        try:
            filename_col = int(self.filename_col_var.get())
        except (tk.TclError, ValueError):
            return

        self.status_col_var.set(min(filename_col + 1, 16384))

    def on_auto_status_toggle(self) -> None:
        if self.auto_status_col_var.get():
            self.on_filename_col_changed()
            self.status_col_spinbox.state(["disabled"])
        else:
            self.status_col_spinbox.state(["!disabled"])

        self._save_settings()

    def _settings_path(self) -> Path:
        appdata = os.getenv("APPDATA")
        if appdata:
            folder = Path(appdata) / "OLEDMeasurementHelper"
        else:
            folder = Path.home() / ".oled_measurement_helper"

        folder.mkdir(parents=True, exist_ok=True)
        return folder / "settings.txt"

    def _load_settings(self) -> None:
        path = self._settings_path()
        if not path.exists():
            return

        try:
            values = {}
            for line in path.read_text(encoding="utf-8").splitlines():
                if "=" not in line:
                    continue
                key, value = line.split("=", 1)
                values[key.strip()] = value.strip()

            self.folder_var.set(values.get("folder", ""))
            self.filename_col_var.set(int(values.get("filename_col", "1")))
            self.status_col_var.set(int(values.get("status_col", "2")))
            self.auto_status_col_var.set(
                values.get("auto_status_col", "1") == "1"
            )
            self.start_row_var.set(int(values.get("start_row", "2")))
            self.extensions_var.set(
                values.get("extensions", DEFAULT_EXTENSIONS)
            )
        except Exception:
            pass

    def _save_settings(self) -> None:
        path = self._settings_path()
        text = "\n".join(
            [
                f"folder={self.folder_var.get().strip()}",
                f"filename_col={self.filename_col_var.get()}",
                f"status_col={self.status_col_var.get()}",
                f"auto_status_col={1 if self.auto_status_col_var.get() else 0}",
                f"start_row={self.start_row_var.get()}",
                f"extensions={self.extensions_var.get().strip()}",
            ]
        )
        path.write_text(text, encoding="utf-8")

    def set_connection_status(
        self,
        connected: bool,
        detail: str,
    ) -> None:
        if connected:
            color = "#2e7d32"
            status = "연결됨"
        else:
            color = "#c62828"
            status = "연결 안 됨"

        self.connection_indicator.config(fg=color)
        self.connection_status_label.config(fg=color)
        self.connection_text_var.set(status)
        self.connection_detail_var.set(detail)

    def connect_excel(self) -> bool:
        try:
            pythoncom.CoInitialize()

            self.excel = win32com.client.GetActiveObject(
                "Excel.Application"
            )
            self.workbook = self.excel.ActiveWorkbook

            if self.workbook is None:
                raise RuntimeError("활성 통합 문서가 없습니다.")

            # CSV가 아닌 관리용 Excel 파일만 연결합니다.
            workbook_name = str(self.workbook.Name)
            if workbook_name.lower().endswith(".csv"):
                raise RuntimeError(
                    "현재 활성 파일이 CSV입니다. "
                    "샘플 목록이 있는 Excel 통합 문서를 선택한 뒤 다시 연결하세요."
                )

            self.sheet = self.workbook.ActiveSheet
            self.user_connected = True

            self.set_connection_status(
                True,
                f"{self.workbook.Name} / {self.sheet.Name}",
            )
            return True

        except Exception as exc:
            self.excel = None
            self.workbook = None
            self.sheet = None
            self.user_connected = False

            self.set_connection_status(False, str(exc))
            return False

    def verify_excel_connection(self) -> bool:
        if not self.user_connected:
            return False

        try:
            if self.excel is None or self.workbook is None:
                return False

            _ = self.workbook.Name
            _ = self.sheet.Name

            # 연결 당시의 통합 문서와 시트를 그대로 유지합니다.
            # 측정 CSV가 Excel에서 열려도 관리 시트로 전환하지 않습니다.
            self.set_connection_status(
                True,
                f"{self.workbook.Name} / {self.sheet.Name}",
            )
            return True

        except Exception:
            self.excel = None
            self.workbook = None
            self.sheet = None
            self.user_connected = False

            self.set_connection_status(
                False,
                "Excel 연결이 끊어졌습니다. 연결 버튼을 다시 누르세요.",
            )
            return False

    def selected_extension(self) -> str:
        """사용 확장자를 '.csv' 형태로 반환합니다."""
        extension = self.extensions_var.get().strip().lower()

        if not extension:
            return ""

        if "," in extension or ";" in extension:
            raise RuntimeError(
                "사용 확장자는 하나만 입력하세요. 예: .csv"
            )

        if not extension.startswith("."):
            extension = "." + extension

        return extension

    def excel_name_key(self, name: str, extension: str) -> str:
        """엑셀 목록 이름의 비교 키를 반환합니다."""
        return normalize_filename(name)

    def saved_file_key(self, name: str, extension: str) -> str:
        """저장 파일 이름의 비교 키를 반환합니다."""
        if extension:
            return normalize_stem(name)

        return normalize_filename(name)

    def get_saved_files(self) -> tuple[set[str], Optional[str]]:
        folder_text = self.folder_var.get().strip()

        if not folder_text:
            raise RuntimeError("측정 폴더를 선택하세요.")

        folder = Path(folder_text)

        if not folder.exists() or not folder.is_dir():
            raise RuntimeError("측정 폴더가 존재하지 않습니다.")

        extension = self.selected_extension()
        saved_keys: set[str] = set()

        newest_name: Optional[str] = None
        newest_time = -1.0

        # 측정 프로그램이 날짜별 하위 폴더를 만드는 경우도 감지합니다.
        for entry in folder.rglob("*"):
            try:
                if not entry.is_file():
                    continue

                if extension and entry.suffix.lower() != extension:
                    continue

                saved_keys.add(self.saved_file_key(entry.name, extension))

                modified_time = entry.stat().st_mtime
                if modified_time > newest_time:
                    newest_time = modified_time
                    try:
                        newest_name = str(entry.relative_to(folder))
                    except ValueError:
                        newest_name = entry.name

            except OSError:
                continue

        return saved_keys, newest_name

    def get_last_row(self) -> int:
        filename_col = self.filename_col_var.get()
        start_row = self.start_row_var.get()

        xl_up = -4162
        last_excel_row = self.sheet.Rows.Count

        last_row = self.sheet.Cells(
            last_excel_row,
            filename_col,
        ).End(xl_up).Row

        return max(start_row, int(last_row))

    def read_names(self) -> list[tuple[int, str]]:
        start_row = self.start_row_var.get()
        filename_col = self.filename_col_var.get()
        last_row = self.get_last_row()

        values = self.sheet.Range(
            self.sheet.Cells(start_row, filename_col),
            self.sheet.Cells(last_row, filename_col),
        ).Value

        rows: list[tuple[int, str]] = []

        if start_row == last_row:
            values = ((values,),)

        for offset, row_value in enumerate(values):
            value = row_value[0] if isinstance(row_value, tuple) else row_value
            name = normalize_text(value)

            if name:
                rows.append((start_row + offset, name))

        return rows

    def write_statuses(
        self,
        rows: list[tuple[int, str]],
        saved_stems: set[str],
    ) -> tuple[int, list[str]]:
        status_col = self.status_col_var.get()

        completed_count = 0
        pending_names: list[str] = []

        if not rows:
            return completed_count, pending_names

        start_row = rows[0][0]
        end_row = rows[-1][0]

        current_values = self.sheet.Range(
            self.sheet.Cells(start_row, status_col),
            self.sheet.Cells(end_row, status_col),
        ).Value

        if start_row == end_row:
            current_values = ((current_values,),)

        output = []
        row_lookup = {row: name for row, name in rows}
        extension = self.selected_extension()

        for row_num in range(start_row, end_row + 1):
            current_value = current_values[row_num - start_row][0]
            sample_name = row_lookup.get(row_num, "")

            if not sample_name:
                output.append((current_value,))
                continue

            is_done = self.excel_name_key(sample_name, extension) in saved_stems

            if is_done:
                completed_count += 1
                output.append(("측정 완료",))
            else:
                pending_names.append(sample_name)

                if current_value == "측정 완료":
                    output.append((None,))
                else:
                    output.append((current_value,))

        self.sheet.Range(
            self.sheet.Cells(start_row, status_col),
            self.sheet.Cells(end_row, status_col),
        ).Value = tuple(output)

        return completed_count, pending_names

    def update_excel_and_clipboard(self) -> None:
        filename_col = self.filename_col_var.get()
        status_col = self.status_col_var.get()

        if status_col <= filename_col:
            self.state_var.set("설정 오류")
            self.footer_var.set(
                "상태 열은 파일 이름 열보다 오른쪽에 있어야 합니다."
            )
            return

        if not self.user_connected:
            self.footer_var.set(
                "엑셀에 연결하려면 '현재 엑셀에 연결' 버튼을 누르세요."
            )
            return

        if not self.verify_excel_connection():
            return

        try:
            saved_stems, newest_name = self.get_saved_files()
            rows = self.read_names()

            completed_count, pending_names = self.write_statuses(
                rows,
                saved_stems,
            )

            total_count = len(rows)

            current_name = pending_names[0] if pending_names else ""
            next_name = (
                pending_names[1]
                if len(pending_names) > 1
                else ""
            )

            if (
                current_name
                and current_name != self.previous_clipboard
            ):
                pyperclip.copy(current_name)
                self.previous_clipboard = current_name

            clipboard_value = normalize_text(pyperclip.paste())

            valid_names = {
                name.casefold(): name for _, name in rows
            }

            if clipboard_value.casefold() in valid_names:
                self.clipboard_var.set(clipboard_value)
            else:
                self.clipboard_var.set("파일 이름 아님")

            self.current_var.set(
                current_name or "모든 측정 완료"
            )
            self.next_var.set(next_name or "-")
            self.saved_var.set(newest_name or "-")

            # 진행률은 완료 개수와 전체 파일 개수를 직접 사용합니다.
            # 퍼센트 반올림 때문에 막대가 부정확하게 보이는 문제를 방지합니다.
            completed_count = total_count - len(pending_names)
            percent = (
                completed_count / total_count * 100
                if total_count
                else 0.0
            )

            self.progress.configure(
                maximum=max(total_count, 1),
                value=completed_count,
            )
            self.progress_text_var.set(
                f"{completed_count} / {total_count} ({percent:.1f}%)"
            )
            self.root.update_idletasks()

            if not rows:
                self.progress.configure(maximum=1, value=0)
                self.progress_text_var.set("0 / 0 (0.0%)")
                self.state_var.set("파일 이름 없음")
            elif completed_count == total_count:
                self.state_var.set("모든 측정 완료")
            else:
                self.state_var.set("측정 대기")

            extension = self.selected_extension()
            matched_names = [
                name for _, name in rows
                if self.excel_name_key(name, extension) in saved_stems
            ]
            if matched_names:
                self.footer_var.set(
                    f"동기화 정상 · 감지 파일 {len(saved_stems)}개 · "
                    f"일치 {len(matched_names)}개 · 최근: {newest_name or '-'}"
                )
            elif newest_name and rows:
                self.footer_var.set(
                    f"파일은 감지했지만 설정 규칙에 따라 이름이 일치하지 않습니다. "
                    f"최근 파일: {newest_name} / 현재 대상: {rows[0][1]}"
                )
            else:
                self.footer_var.set(
                    f"저장 파일을 찾지 못했습니다. 확인 폴더: {self.folder_var.get().strip()}"
                )

            self._save_settings()

        except Exception as exc:
            self.footer_var.set(f"오류: {exc}")
            self.state_var.set("오류")

    def copy_current_target(self) -> None:
        name = self.current_var.get().strip()

        if name and name != "모든 측정 완료":
            pyperclip.copy(name)
            self.previous_clipboard = name
            self.clipboard_var.set(name)

    def refresh_all(self) -> None:
        if not self.running:
            return

        self.update_excel_and_clipboard()
        self.root.after(CHECK_INTERVAL_MS, self.refresh_all)

    def on_close(self) -> None:
        self.running = False
        self._save_settings()

        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    MeasurementHelper(root)
    root.mainloop()


if __name__ == "__main__":
    main()
