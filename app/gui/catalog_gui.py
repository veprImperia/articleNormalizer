from __future__ import annotations

import json
import os
import sys
import traceback
import threading
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ---------------------------------------------------------
# IMPORTS: поправь под свою структуру проекта
# ---------------------------------------------------------
# Вариант 1: если gui-файл лежит рядом с модулями
# from docx_manager import DocxParser
# from merge_selector import MergeSelector
# from structure_restorer import StructureRestorer
# from merger import Merger
# from formatter import CatalogFormatter

# Вариант 2: если у тебя структура app/core/...
from app.core.docx_manager import DocxParser, DocxValidator
from app.core.merge_selector import MergeSelector
from app.core.structure_restorer import StructureRestorer
from app.core.merger import Merger
from app.core.formatter import CatalogFormatter
from app.io.io_manager import InputCollector


def normalize_validation_report(report: dict | None) -> dict:
    if report is None:
        return {
            "status": "error",
            "issues": ["validation_none"],
            "optional_miss": [],
        }

    compact = {
        "status": report.get("status", "error"),
        "issues": list(dict.fromkeys(report.get("issues", []))),
        "optional_miss": list(dict.fromkeys(report.get("optional_miss", []))),
    }

    if "_doc_kind" in report:
        compact["_doc_kind"] = report["_doc_kind"]

    return compact


def has_minimal_merge_blocks(final_blocks: dict | None) -> bool:
    if not final_blocks:
        return False

    required = (
        "title_ru",
        "authors_ru_block",
        "abstract_ru",
        "keywords_ru",
    )
    return all(final_blocks.get(k) not in (None, []) for k in required)


def compute_merge_score(file_name: str, final_blocks: dict | None) -> int:
    if not final_blocks:
        return 0

    score = 0

    if final_blocks.get("title_ru") not in (None, []):
        score += 5
    if final_blocks.get("authors_ru_block") not in (None, []):
        score += 5
    if final_blocks.get("abstract_ru") not in (None, []):
        score += 5
    if final_blocks.get("keywords_ru") not in (None, []):
        score += 5

    if final_blocks.get("body_block") not in (None, []):
        score += 2
    if final_blocks.get("reference_block") not in (None, []):
        score += 1

    if final_blocks.get("title_en") not in (None, []):
        score += 1
    if final_blocks.get("authors_en_block") not in (None, []):
        score += 1
    if final_blocks.get("abstract_en") not in (None, []):
        score += 1
    if final_blocks.get("keywords_en") not in (None, []):
        score += 1

    name_low = file_name.lower()

    if "тезис" in name_low:
        score += 3
    if "доклад" in name_low:
        score += 2
    if "аннотац" in name_low or "abstract" in name_low or "annotation" in name_low:
        score += 1

    if "заявка" in name_low:
        score -= 100
    if "summary" in name_low:
        score -= 100
    if "template" in name_low or "шаблон" in name_low:
        score -= 50

    if "(1)" in name_low or "(2)" in name_low or "v2" in name_low or "v3" in name_low:
        score -= 1

    return score

def split_report_by_status(report_dict: dict[str, dict]) -> dict[str, dict[str, dict]]:
    buckets = {
        "ok": {},
        "partial": {},
        "invalid": {},
        "ignored": {},
        "error": {},
        "unknown": {},
    }

    for file_name, report in report_dict.items():
        status = report.get("status", "unknown")
        if status not in buckets:
            status = "unknown"
        buckets[status][file_name] = report

    return buckets


def looks_like_application(file_name: str) -> bool:
    name = file_name.lower()
    markers = (
        "заявка",
        "application",
        "participant",
        "участник",
        "регистрац",
        "анкета",
        "заключение",
        "эксперт",
    )
    return any(m in name for m in markers)


def looks_like_english_only(file_name: str, report: dict) -> bool:
    name = file_name.lower()
    issues = set(report.get("issues", []))

    if "англ" in name or "_en" in name or "annotation" in name or "abstract" in name:
        if "missing_title_ru" in issues and "missing_authors_ru_block" in issues:
            return True

    if (
        "missing_title_ru" in issues
        and "missing_authors_ru_block" in issues
        and "missing_abstract_ru" in issues
    ):
        return True

    return False


def looks_like_header_bug_candidate(report: dict) -> bool:
    issues = set(report.get("issues", []))
    return (
        "missing_title_ru" in issues
        and "missing_authors_ru_block" in issues
        and "missing_abstract_ru" not in issues
    )


def looks_like_missing_abstract_candidate(report: dict) -> bool:
    issues = set(report.get("issues", []))
    return "missing_abstract_ru" in issues and "missing_title_ru" not in issues


def extract_duplicate_stem(file_name: str) -> str:
    name = file_name.lower()

    trash = [
        ".docx",
        "аннотация",
        "annotation",
        "abstract",
        "тезисы",
        "tezisy",
        "thesis",
        "doclad",
        "доклад",
        "заявка",
        "application",
        "ред",
        "испр",
        "(1)",
        "(2)",
        " v2",
        " v3",
        "_",
        "-",
        ".",
        ",",
    ]

    for item in trash:
        name = name.replace(item, " ")

    name = " ".join(name.split())
    return name.strip()


def build_duplicate_groups(report_dict: dict[str, dict]) -> dict[str, list[str]]:
    groups: dict[str, list[str]] = {}

    for file_name in report_dict:
        stem = extract_duplicate_stem(file_name)
        if len(stem) < 4:
            continue
        groups.setdefault(stem, []).append(file_name)

    return {
        stem: sorted(files)
        for stem, files in groups.items()
        if len(files) > 1
    }


def split_problem_buckets(report_dict: dict[str, dict]) -> dict[str, dict[str, dict]]:
    applications = {}
    english_only = {}
    header_bug = {}
    missing_abstract = {}

    for file_name, report in report_dict.items():
        status = report.get("status")

        if status not in {"partial", "invalid", "ignored"}:
            continue

        if looks_like_application(file_name):
            applications[file_name] = report
            continue

        if looks_like_english_only(file_name, report):
            english_only[file_name] = report
            continue

        if looks_like_header_bug_candidate(report):
            header_bug[file_name] = report
            continue

        if looks_like_missing_abstract_candidate(report):
            missing_abstract[file_name] = report
            continue

    return {
        "applications": applications,
        "english_only": english_only,
        "header_bug_candidates": header_bug,
        "missing_abstract_candidates": missing_abstract,
    }


def build_duplicate_key(file_name: str) -> str:
    name = file_name.lower()

    trash = [
        ".docx",
        "аннотация",
        "annotation",
        "abstract",
        "тезисы",
        "тезис",
        "tezisy",
        "thesis",
        "doclad",
        "доклад",
        "заявка",
        "application",
        "summary",
        "template",
        "шаблон",
        "ред",
        "испр",
        "(1)",
        "(2)",
        "v2",
        "v3",
        "_",
        "-",
        ".",
        ",",
    ]

    for item in trash:
        name = name.replace(item, " ")

    name = " ".join(name.split())
    return name.strip()


def build_debug_payload(parser: DocxParser, validator: DocxValidator, file_path: Path) -> dict:
    payload = {
        "file_name": file_path.name,
        "file_path": str(file_path),
        "error": None,
        "final_blocks": None,
        "validation": None,
    }

    try:
        final_blocks = parser.get_parse_data(file_path)
        payload["final_blocks"] = final_blocks

        validation = validator.validate_doc_struct(final_blocks)
        payload["validation"] = validation
        return payload

    except Exception as exc:
        payload["error"] = f"{type(exc).__name__}: {exc}"
        payload["validation"] = {
            "status": "error",
            "issues": [f"exception:{type(exc).__name__}"],
            "optional_miss": [],
        }
        return payload


def build_report_dict(results: dict[str, dict]) -> dict[str, dict]:
    report = {}

    for file_name, payload in results.items():
        validation = normalize_validation_report(payload.get("validation"))
        final_blocks = payload.get("final_blocks")

        report[file_name] = {
            **validation,
            "merge_ready_minimal": has_minimal_merge_blocks(final_blocks),
            "merge_score": compute_merge_score(file_name, final_blocks),
            "duplicate_key": build_duplicate_key(file_name),
        }

    return report


class CatalogPipelineGUI(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Catalog Builder")
        self.geometry("1200x820")
        self.minsize(1050, 720)

        self.collected_dir_var = tk.StringVar()
        self.report_json_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        self.input_root_var = tk.StringVar()
        self.use_input_collector_var = tk.BooleanVar(value=True)
        self.use_formatter_var = tk.BooleanVar(value=True)

        self.summary_vars = {
            "ok": tk.StringVar(value="0"),
            "partial": tk.StringVar(value="0"),
            "invalid": tk.StringVar(value="0"),
            "ignored": tk.StringVar(value="0"),
            "error": tk.StringVar(value="0"),
            "unknown": tk.StringVar(value="0"),
        }

        self._build_ui()
        self._log("GUI готов.")

    # ---------------------------------------------------------
    # UI
    # ---------------------------------------------------------

    def _build_ui(self):
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)

        title = ttk.Label(
            root,
            text="Склейка каталога тезисов",
            font=("Segoe UI", 14, "bold"),
        )
        title.pack(anchor="w", pady=(0, 10))

        # ---------------- paths ----------------
        path_box = ttk.LabelFrame(root, text="Пути", padding=10)
        path_box.pack(fill="x", pady=(0, 10))
        

        self._add_path_row(
            parent=path_box,
            label="Input root",
            var=self.input_root_var,
            browse_cmd=self._choose_input_root,
            row=0,
        )

        self._add_path_row(
            parent=path_box,
            label="Collected dir",
            var=self.collected_dir_var,
            browse_cmd=self._choose_collected_dir,
            row=1,
        )
        self._add_path_row(
            parent=path_box,
            label="Report json",
            var=self.report_json_var,
            browse_cmd=self._choose_report_json,
            row=2,
        )
        self._add_path_row(
            parent=path_box,
            label="Output dir",
            var=self.output_dir_var,
            browse_cmd=self._choose_output_dir,
            row=3,
        )

        # ---------------- options ----------------
        extra_box = ttk.LabelFrame(root, text="Опции", padding=10)
        extra_box.pack(fill="x", pady=(0, 10))

        ttk.Label(
            extra_box,
            text="Plenary titles (по одному на строку, опционально)"
        ).grid(row=0, column=0, sticky="w")

        self.plenary_text = tk.Text(extra_box, height=5, wrap="word")
        self.plenary_text.grid(row=1, column=0, columnspan=4, sticky="nsew", pady=(4, 8))

        ttk.Checkbutton(
            extra_box,
            text="Пост-форматирование итогового docx",
            variable=self.use_formatter_var,
        ).grid(row=2, column=0, sticky="w")

        ttk.Checkbutton(
            extra_box,
            text="Сначала собрать docx из сырой папки через InputCollector",
            variable=self.use_input_collector_var,
        ).grid(row=3, column=0, sticky="w", pady=(8, 0))

        extra_box.columnconfigure(0, weight=1)

        # ---------------- run buttons ----------------
        btn_box = ttk.Frame(root)
        btn_box.pack(fill="x", pady=(0, 10))

        self.run_btn = ttk.Button(
            btn_box,
            text="Запустить pipeline",
            command=self._start_pipeline_thread,
        )
        self.run_btn.pack(side="left")

        ttk.Button(
            btn_box,
            text="Открыть output dir",
            command=self._open_output_dir,
        ).pack(side="left", padx=8)

        ttk.Button(
            btn_box,
            text="Очистить лог",
            command=self._clear_log,
        ).pack(side="left")

        

        # ---------------- summary ----------------
        summary_box = ttk.LabelFrame(root, text="Сводка", padding=10)
        summary_box.pack(fill="x", pady=(0, 10))

        self._add_summary_label(summary_box, "OK", "ok", 0)
        self._add_summary_label(summary_box, "Partial", "partial", 1)
        self._add_summary_label(summary_box, "Invalid", "invalid", 2)
        self._add_summary_label(summary_box, "Ignored", "ignored", 3)
        self._add_summary_label(summary_box, "Error", "error", 4)
        self._add_summary_label(summary_box, "Unknown", "unknown", 5)

        # ---------------- reports ----------------
        reports_box = ttk.LabelFrame(root, text="Отчёты и артефакты", padding=10)
        reports_box.pack(fill="x", pady=(0, 10))


        ttk.Button(
            reports_box,
            text="Показать collector report",
            command=lambda: self._show_named_report("collector_report.txt"),
        ).grid(row=4, column=0, padx=4, pady=4, sticky="ew")

        ttk.Button(
            reports_box,
            text="Open quarantine",
            command=lambda: self._open_named_subdir("input_quarantine"),
        ).grid(row=4, column=1, padx=4, pady=4, sticky="ew")

        ttk.Button(
            reports_box,
            text="Open collector protocol",
            command=lambda: self._show_named_report("collector_protocol.json"),
        ).grid(row=4, column=2, padx=4, pady=4, sticky="ew")

        ttk.Button(
            reports_box,
            text="Показать summary",
            command=lambda: self._show_named_report("summary_doc.json"),
        ).grid(row=0, column=0, padx=4, pady=4, sticky="ew")

        ttk.Button(
            reports_box,
            text="Показать report",
            command=lambda: self._show_named_report("report_doc.json"),
        ).grid(row=0, column=1, padx=4, pady=4, sticky="ew")

        ttk.Button(
            reports_box,
            text="Показать debug log",
            command=lambda: self._show_named_report("debug_parser.txt"),
        ).grid(row=0, column=2, padx=4, pady=4, sticky="ew")

        ttk.Button(
            reports_box,
            text="Показать merge report",
            command=lambda: self._show_named_report("merge_report.txt"),
        ).grid(row=0, column=3, padx=4, pady=4, sticky="ew")

        ttk.Button(
            reports_box,
            text="Показать restorer report",
            command=lambda: self._show_named_report("restorer_report.txt"),
        ).grid(row=1, column=0, padx=4, pady=4, sticky="ew")

        ttk.Button(
            reports_box,
            text="Header bugs",
            command=lambda: self._show_named_report("header_bug_candidates.json"),
        ).grid(row=1, column=1, padx=4, pady=4, sticky="ew")

        ttk.Button(
            reports_box,
            text="Missing abstract",
            command=lambda: self._show_named_report("missing_abstract_candidates.json"),
        ).grid(row=1, column=2, padx=4, pady=4, sticky="ew")

        ttk.Button(
            reports_box,
            text="Duplicates",
            command=lambda: self._show_named_report("likely_duplicates.json"),
        ).grid(row=1, column=3, padx=4, pady=4, sticky="ew")

        ttk.Button(
            reports_box,
            text="English only",
            command=lambda: self._show_named_report("english_only_doc.json"),
        ).grid(row=2, column=0, padx=4, pady=4, sticky="ew")

        ttk.Button(
            reports_box,
            text="Runtime errors",
            command=lambda: self._show_named_report("runtime_errors_detail.json"),
        ).grid(row=2, column=1, padx=4, pady=4, sticky="ew")

        ttk.Button(
            reports_box,
            text="Open summary externally",
            command=lambda: self._open_named_report("summary_doc.json"),
        ).grid(row=2, column=2, padx=4, pady=4, sticky="ew")

        ttk.Button(
            reports_box,
            text="Open debug externally",
            command=lambda: self._open_named_report("debug_parser.txt"),
        ).grid(row=2, column=3, padx=4, pady=4, sticky="ew")

        ttk.Button(
            reports_box,
            text="Open merged docx",
            command=lambda: self._open_named_report("catalog_merged.docx"),
        ).grid(row=3, column=0, padx=4, pady=4, sticky="ew")

        ttk.Button(
            reports_box,
            text="Open formatted docx",
            command=lambda: self._open_named_report("catalog_formatted.docx"),
        ).grid(row=3, column=1, padx=4, pady=4, sticky="ew")

        ttk.Button(
            reports_box,
            text="Open for_merge",
            command=lambda: self._open_named_subdir("for_merge"),
        ).grid(row=3, column=2, padx=4, pady=4, sticky="ew")

        ttk.Button(
            reports_box,
            text="Open manual_review",
            command=lambda: self._open_named_subdir("manual_review"),
        ).grid(row=3, column=3, padx=4, pady=4, sticky="ew")

        for i in range(4):
            reports_box.columnconfigure(i, weight=1)

        # ---------------- log / viewer ----------------
        log_box = ttk.LabelFrame(root, text="Viewer / Лог", padding=10)
        log_box.pack(fill="both", expand=True)

        self.log_text = tk.Text(log_box, wrap="word", height=22)
        self.log_text.pack(side="left", fill="both", expand=True)

        scroll = ttk.Scrollbar(log_box, command=self.log_text.yview)
        scroll.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scroll.set)

    def _add_path_row(self, parent, label, var, browse_cmd, row: int):
        ttk.Label(parent, text=label, width=14).grid(row=row, column=0, sticky="w", pady=4)
        entry = ttk.Entry(parent, textvariable=var)
        entry.grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        ttk.Button(parent, text="Выбрать", command=browse_cmd).grid(row=row, column=2, sticky="ew", pady=4)
        parent.columnconfigure(1, weight=1)

    def _add_summary_label(self, parent, title: str, key: str, column: int):
        frame = ttk.Frame(parent)
        frame.grid(row=0, column=column, padx=8, sticky="w")
        ttk.Label(frame, text=f"{title}: ", font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Label(frame, textvariable=self.summary_vars[key]).pack(side="left")

    # ---------------------------------------------------------
    # browse
    # ---------------------------------------------------------

    def _choose_input_root(self):
        path = filedialog.askdirectory(title="Выбери сырую папку input_root")
        if path:
            self.input_root_var.set(path)

    def _choose_collected_dir(self):
        path = filedialog.askdirectory(title="Выбери папку collected_dir")
        if path:
            self.collected_dir_var.set(path)

    def _choose_report_json(self):
        path = filedialog.askopenfilename(
            title="Выбери report.json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if path:
            self.report_json_var.set(path)

    def _choose_output_dir(self):
        path = filedialog.askdirectory(title="Выбери output dir")
        if path:
            self.output_dir_var.set(path)

    # ---------------------------------------------------------
    # log / viewer
    # ---------------------------------------------------------

    def _log(self, msg: str):
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.update_idletasks()

    def _clear_log(self):
        self.log_text.delete("1.0", "end")

    def _show_file_in_log(self, path: Path):
        if not path.exists():
            messagebox.showinfo("Info", f"Файл не найден:\n{path}")
            return

        try:
            if path.suffix.lower() == ".json":
                data = json.loads(path.read_text(encoding="utf-8"))
                text = json.dumps(data, indent=2, ensure_ascii=False)
            else:
                text = path.read_text(encoding="utf-8", errors="replace")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось прочитать файл:\n{e}")
            return

        self._clear_log()
        self._log(f"=== {path.name} ===\n")
        self._log(text)

    # ---------------------------------------------------------
    # open external
    # ---------------------------------------------------------

    def _open_file(self, path: Path):
        if not path.exists():
            messagebox.showinfo("Info", f"Файл не найден:\n{path}")
            return

        try:
            if sys.platform.startswith("win"):
                os.startfile(str(path))
            elif sys.platform == "darwin":
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}"')
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть файл:\n{e}")

    def _open_output_dir(self):
        path = self.output_dir_var.get().strip()
        if not path:
            messagebox.showinfo("Info", "Сначала выбери output dir")
            return

        p = Path(path)
        if not p.exists():
            messagebox.showinfo("Info", "Папка output dir пока не существует")
            return

        self._open_file(p)

    def _report_path(self, filename: str) -> Path:
        output_dir = Path(self.output_dir_var.get().strip())
        return output_dir / filename

    def _show_named_report(self, filename: str):
        path = self._report_path(filename)
        self._show_file_in_log(path)

    def _open_named_report(self, filename: str):
        path = self._report_path(filename)
        self._open_file(path)

    def _open_named_subdir(self, dirname: str):
        output_dir = Path(self.output_dir_var.get().strip())
        self._open_file(output_dir / dirname)

    # ---------------------------------------------------------
    # pipeline run
    # ---------------------------------------------------------

    def _start_pipeline_thread(self):
        if not self._validate_inputs():
            return

        self.run_btn.config(state="disabled")
        worker = threading.Thread(target=self._run_pipeline_safe, daemon=True)
        worker.start()

    def _run_pipeline_safe(self):
        try:
            self._run_pipeline()
            self.after(0, lambda: messagebox.showinfo("Готово", "Pipeline завершён успешно"))
        except Exception as e:
            err = "".join(traceback.format_exception(type(e), e, e.__traceback__))
            self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
            self.after(0, lambda: self._log("\n[ERROR]\n" + err))
        finally:
            self.after(0, lambda: self.run_btn.config(state="normal"))

    def _run_pipeline(self):
        collected_dir = Path(self.collected_dir_var.get()).resolve()
        output_dir = Path(self.output_dir_var.get()).resolve()
        input_root = Path(self.input_root_var.get()).resolve() if self.input_root_var.get().strip() else None

        output_dir.mkdir(parents=True, exist_ok=True)

        # --- paths ---
        quarantine_dir = output_dir / "input_quarantine"

        report_json = output_dir / "report_doc.json"
        summary_json = output_dir / "summary_doc.json"
        debug_log_txt = output_dir / "debug_parser.txt"

        selector_protocol_json = output_dir / "merge_protocol.json"
        selector_report_txt = output_dir / "merge_report.txt"

        for_merge_dir = output_dir / "for_merge"
        manual_review_dir = output_dir / "manual_review"

        restored_blocks_json = output_dir / "restored_blocks.json"
        restorer_protocol_json = output_dir / "restorer_protocol.json"
        restorer_report_txt = output_dir / "restorer_report.txt"

        merged_docx = output_dir / "catalog_merged.docx"
        formatted_docx = output_dir / "catalog_formatted.docx"

        plenary_titles = self._read_plenary_titles()

        self._clear_log()
        self._log(f"Plenary titles loaded: {len(plenary_titles)}")
        for idx, t in enumerate(plenary_titles, start=1):
            self._log(f"  {idx}. {t}")
            self._log("")

        # =========================================================
        # STEP 0: InputCollector
        # =========================================================
        if self.use_input_collector_var.get():
            self._log("=== STEP 0: InputCollector ===")

            collector = InputCollector(
                input_root=input_root,
                collected_dir=collected_dir,
                quarantine_dir=quarantine_dir,
            )
            collector.collect_files()
            collector.save_protocol(output_dir / "collector_protocol.json")
            collector.save_human_report(output_dir / "collector_report.txt")

            self._log(f"Accepted: {len(collector.protocol.get('accepted_packages', []))}")
            self._log(f"Quarantined: {len(collector.protocol.get('quarantined_packages', []))}")
            self._log(f"Copied docx: {len(collector.protocol.get('copied_docx', []))}")
            self._log("")

        # =========================================================
        # STEP 1: Parser → report_doc.json
        # =========================================================
        self._log("=== STEP 1: Parsing docx → report ===")

        parser = DocxParser()
        validator = DocxValidator()

        results = {}
        docx_files = list(collected_dir.rglob("*.docx"))

        if not docx_files:
            self._log("[WARNING] В collected_dir нет .docx файлов")

        for file in docx_files:
            payload = build_debug_payload(parser, validator, file)
            results[file.name] = payload

        report = build_report_dict(results)

        summary = {
            "by_status": {
                "ok": 0,
                "partial": 0,
                "invalid": 0,
                "ignored": 0,
                "error": 0,
                "unknown": 0,
            }
        }

        debug_lines = []

        for file_name, payload in results.items():
            validation = normalize_validation_report(payload.get("validation"))
            status = validation.get("status", "unknown")

            if status not in summary["by_status"]:
                status = "unknown"

            summary["by_status"][status] += 1

            if payload.get("error"):
                debug_lines.append(f"{file_name} :: ERROR :: {payload['error']}")

        report_json.write_text(
            json.dumps(report, indent=2, ensure_ascii=False),
            encoding="utf-8"
        )

        summary_json.write_text(
            json.dumps(summary, indent=2, ensure_ascii=False),
            encoding="utf-8"
        )
        self._update_summary_vars(summary["by_status"])

        debug_log_txt.write_text("\n".join(debug_lines), encoding="utf-8")

        # --- extra reports ---
        buckets = split_report_by_status(report)
        (output_dir / "ok_doc.json").write_text(
            json.dumps(buckets["ok"], indent=2, ensure_ascii=False), encoding="utf-8"
        )
        (output_dir / "partial_doc.json").write_text(
            json.dumps(buckets["partial"], indent=2, ensure_ascii=False), encoding="utf-8"
        )
        (output_dir / "invalid_doc.json").write_text(
            json.dumps(buckets["invalid"], indent=2, ensure_ascii=False), encoding="utf-8"
        )
        (output_dir / "ignored_doc.json").write_text(
            json.dumps(buckets["ignored"], indent=2, ensure_ascii=False), encoding="utf-8"
        )

        problem_buckets = split_problem_buckets(report)
        (output_dir / "applications_doc.json").write_text(
            json.dumps(problem_buckets["applications"], indent=2, ensure_ascii=False), encoding="utf-8"
        )
        (output_dir / "english_only_doc.json").write_text(
            json.dumps(problem_buckets["english_only"], indent=2, ensure_ascii=False), encoding="utf-8"
        )
        (output_dir / "header_bug_candidates.json").write_text(
            json.dumps(problem_buckets["header_bug_candidates"], indent=2, ensure_ascii=False), encoding="utf-8"
        )
        (output_dir / "missing_abstract_candidates.json").write_text(
            json.dumps(problem_buckets["missing_abstract_candidates"], indent=2, ensure_ascii=False), encoding="utf-8"
        )

        duplicate_groups = build_duplicate_groups(report)
        (output_dir / "likely_duplicates.json").write_text(
            json.dumps(duplicate_groups, indent=2, ensure_ascii=False), encoding="utf-8"
        )

        runtime_errors = {
            file_name: {
                "status": rep.get("status"),
                "issues": rep.get("issues", []),
            }
            for file_name, rep in report.items()
            if rep.get("status") == "error"
        }
        (output_dir / "runtime_errors_detail.json").write_text(
            json.dumps(runtime_errors, indent=2, ensure_ascii=False), encoding="utf-8"
        )

        self._log(f"Parsed files: {len(results)}")
        self._log(f"Summary: {summary['by_status']}")
        self._log("")

        # =========================================================
        # STEP 2: MergeSelector
        # =========================================================
        self._log("=== STEP 2: MergeSelector ===")

        selector = MergeSelector(
            collected_dir=collected_dir,
            report_json_path=report_json,
            for_merge_dir=for_merge_dir,
            manual_review_dir=manual_review_dir,
        )

        selector.run()
        selector.save_protocol(selector_protocol_json)
        selector.save_human_report(selector_report_txt)

        self._log(f"Selected: {len(selector.protocol['selected_for_merge'])}")
        self._log(f"Manual: {len(selector.protocol['manual_review'])}")
        self._log(f"Excluded: {len(selector.protocol['excluded'])}")
        self._log("")

        # =========================================================
        # STEP 3: StructureRestorer
        # =========================================================
        self._log("=== STEP 3: StructureRestorer ===")

        restorer = StructureRestorer(parser=parser, for_merge_dir=for_merge_dir)
        blocks = restorer.run()

        restorer.save_blocks(restored_blocks_json, blocks)
        restorer.save_protocol(restorer_protocol_json)
        restorer.save_human_report(restorer_report_txt)

        self._log(f"Restored docs: {len(blocks)}")
        self._log("")

        # =========================================================
        # STEP 4: Merger
        # =========================================================
        self._log("=== STEP 4: Merger ===")

        merger = Merger(
            restored_blocks_json_path=restored_blocks_json,
            output_docx_path=merged_docx,
            plenary_titles=plenary_titles,
        )

        merger.run()

        self._log(f"Merged: {merged_docx}")
        self._log("")

        # =========================================================
        # STEP 5: Formatter (optional)
        # =========================================================
        if self.use_formatter_var.get():
            self._log("=== STEP 5: Formatter ===")

            formatter = CatalogFormatter(
                input_docx_path=merged_docx,
                output_docx_path=formatted_docx,
            )
            formatter.run()

            self._log(f"Formatted: {formatted_docx}")
        else:
            self._log("=== STEP 5: Formatter skipped ===")

        self._log("\n=== DONE ===")
        self._log(f"Output: {output_dir}")

    def _read_plenary_titles(self) -> list[str]:
        raw = self.plenary_text.get("1.0", "end").strip()
        if not raw:
            return []

        titles = []
        for line in raw.splitlines():
            t = line.strip()
            if not t:
                continue

            # убираем внешние кавычки
            if len(t) >= 2 and t[0] == t[-1] and t[0] in {'"', "'", "«", "»"}:
                t = t[1:-1].strip()

            # убираем типографские кавычки по краям, если они перекошены
            t = t.strip('"\''"«»“”„‟")

            if t:
                titles.append(t)

        return titles

    # ---------------------------------------------------------
    # summary
    # ---------------------------------------------------------


    def _update_summary_vars(self, by_status: dict):
        def _apply():
            for key in self.summary_vars:
                self.summary_vars[key].set(str(by_status.get(key, 0)))
            self.update_idletasks()
        self.after(0, _apply)

    def _update_summary_from_existing(self, summary_path: Path):
        if not summary_path.exists():
            return

        try:
            data = json.loads(summary_path.read_text(encoding="utf-8"))
        except Exception:
            return

        by_status = data.get("by_status", {})
        self._update_summary_vars(by_status)

    # ---------------------------------------------------------
    # validation
    # ---------------------------------------------------------

    def _validate_inputs(self) -> bool:
        input_root = self.input_root_var.get().strip()
        collected_dir = self.collected_dir_var.get().strip()
        output_dir = self.output_dir_var.get().strip()

        # report_json теперь НЕ обязательный
        report_json = self.report_json_var.get().strip()

        if self.use_input_collector_var.get():
            if not input_root:
                messagebox.showwarning("Проверка", "Не выбрана папка input_root")
                return False
            if not Path(input_root).exists():
                messagebox.showwarning("Проверка", "input_root не существует")
                return False

        if not collected_dir:
            messagebox.showwarning("Проверка", "Не выбрана папка collected_dir")
            return False
        if not Path(collected_dir).exists():
            messagebox.showwarning("Проверка", "collected_dir не существует")
            return False

        if not output_dir:
            messagebox.showwarning("Проверка", "Не выбрана output dir")
            return False

        return True


