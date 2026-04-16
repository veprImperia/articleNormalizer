# test_docx_parser.py
"""
Диагностический pytest-раннер для пакетной проверки DocxParser.

Что делает:
1. Собирает входные .docx через InputCollector или работает с одним файлом.
2. Для каждого документа строит подробный debug payload:
   - raw_doc_struct
   - clean_struct
   - raw_markers
   - full_map
   - primary_split
   - fallback
   - final_blocks
   - validation
3. Записывает:
   - report_doc.json   -> компактный итог по всем файлам
   - summary_doc.json  -> сводка по статусам
   - debug_parser.txt  -> подробный лог по каждому документу
4. Pytest-тест считается успешным, если не было runtime errors.
   Наличие invalid / partial документов тест НЕ валит.

Зачем:
- чтобы можно было массово гонять парсер через pytest
- чтобы коллега мог быстро понять, где именно ломается разбор
- чтобы потом легко добавлять новые диагностические хуки
"""

from __future__ import annotations
import traceback
import json
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from app.core.docx_manager import DocxParser, DocxValidator
from app.io.io_manager import InputCollector
from app.core.merge_selector import MergeSelector
from app.core.structure_restorer import StructureRestorer
from app.core.merger import Merger
from app.core.formatter import CatalogFormatter

# ============================================================
# CONFIG
# ============================================================

@dataclass(frozen=True)
class RunnerConfig:
    mode: str
    test_docx_path: Path
    input_root: Path
    collected_dir: Path
    quarantine_dir: Path

    for_merge_dir: Path
    manual_review_dir: Path
    merge_protocol_json_path: Path
    merge_report_txt_path: Path

    restored_blocks_json_path: Path
    restorer_protocol_json_path: Path
    restorer_report_txt_path: Path

    merged_output_docx_path: Path
    plenary_titles_json_path: Path

    formatted_output_docx_path: Path

    log_path: Path
    report_json_path: Path
    summary_json_path: Path
    invalid_json_path: Path
    partial_json_path: Path
    ok_json_path: Path
    ignored_json_path: Path

    collector_protocol_json_path: Path
    collector_report_txt_path: Path

    applications_json_path: Path
    english_only_json_path: Path
    header_bug_json_path: Path
    missing_abstract_json_path: Path
    likely_duplicates_json_path: Path
    suspects_protocol_json_path: Path

    suspects: tuple[str, ...]
    log_only_bad: bool = False


CONFIG = RunnerConfig(
    mode="folder",  # "single" | "folder" | "suspects"
    test_docx_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\Пример оформления.docx"
    ),
    input_root=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\row_doc_dir"
    ),
    collected_dir=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\input_collected"
    ),
    for_merge_dir=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\for_merge"
    ),
    manual_review_dir=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\manual_review"
    ),
    merge_protocol_json_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\merge_protocol.json"
    ),
    merge_report_txt_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\merge_report.txt"
    ),
    restored_blocks_json_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\restored_blocks.json"
    ),
    restorer_protocol_json_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\restorer_protocol.json"
    ),
    restorer_report_txt_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\restorer_report.txt"
    ),
    merged_output_docx_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\catalog_merged.docx"
    ),
    plenary_titles_json_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\plenary_titles.json"
    ),
    formatted_output_docx_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\catalog_formatted.docx"
    ),
    log_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\debug_parser.txt"
    ),
    report_json_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\report_doc.json"
    ),
    summary_json_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\summary_doc.json"
    ),
    invalid_json_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\invalid_doc.json"
    ),
    partial_json_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\partial_doc.json"
    ),
    ok_json_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\ok_doc.json"
    ),
    ignored_json_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\ignored_doc.json"
    ),
    applications_json_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\applications_doc.json"
    ),
    english_only_json_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\english_only_doc.json"
    ),
    header_bug_json_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\header_bug_candidates.json"
    ),
    missing_abstract_json_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\missing_abstract_candidates.json"
    ),
    likely_duplicates_json_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\likely_duplicates.json"
    ),
    suspects_protocol_json_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\suspects_protocol.json"
    ),
    quarantine_dir=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\input_quarantine"
    ),
    collector_protocol_json_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\collector_protocol.json"
    ),
    collector_report_txt_path=Path(
        r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\collector_report.txt"
    ),
    log_only_bad = False,
    suspects = (
    "Аннотация-Грибова-ОВ.docx",
    "Аннотация Лобанов FAPM-2025.docx",
    "Аннотация Орёл Н А.docx",
    "Аннотация_FAPM_2025.docx",
    "Аннотация_v3.docx",
    "БеличенкоДА_FAPM-2025_аннотация.docx",
    "БеличенкоДА_FAPM-2025_тезис.docx",
    "БулатовВВ_аннотация доклада Rus and Eng.docx",
    "БулатовВВ_тезисы доклада Rus and Eng.docx",
    "Бычков Р.С._Аннотация.docx",
    "Бычков Р.С._Тезисы.docx",
    "Кутузов аннотация FAPM-2025.docx",
    "Кутузов тезисы FAPM-2025.docx",
    "Тезисы - ГрибоваОВ.docx",
    "Саяпин С.Н..docx",
    "Zubin_Maksimov.docx",
    "АННОТАЦИЯ_FAPM 2025_Зубин.docx",
    "Аннотация_Большенко.docx",
    "Наконечный Хохлов Тезисы_FAPM.docx",
    "Заграничная (аннотация).docx",
    "Annotation Zditovets.docx",
    "Korotkov_Minaeva_Tezisy_FAPM2025_en.docx",
    "АННОТАЦИЯ_англ_FAPM 2025_Зубин.docx",
    "аннотация_Тагировой_FAPM25_англ_яз .docx",
    "Summary.docx",
)

)


# ============================================================
# FILE IO HELPERS
# ============================================================

def ensure_parent_dir(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def clear_text_file(path: Path) -> None:
    ensure_parent_dir(path)
    with open(path, "w", encoding="utf-8") as f:
        f.write("")


def append_text_line(path: Path, text: str = "") -> None:
    ensure_parent_dir(path)
    with open(path, "a", encoding="utf-8") as f:
        f.write(text + "\n")


def save_json(path: Path, data: Any) -> None:
    ensure_parent_dir(path)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def append_json_block(path: Path, title: str, data: Any) -> None:
    ensure_parent_dir(path)
    with open(path, "a", encoding="utf-8") as f:
        f.write(f"{title}:\n")
        json.dump(data, f, indent=2, ensure_ascii=False)
        f.write("\n\n")


# ============================================================
# CORE DATA BUILDERS
# ============================================================

def collect_target_files(config: RunnerConfig) -> list[Path]:
    """
    Возвращает список файлов для прогона согласно mode.
    """
    if config.mode == "single":
        return [config.test_docx_path]

    collector = InputCollector(
        config.input_root,
        config.collected_dir,
        config.quarantine_dir,
    )
    collector.collect_files()
    collector.save_protocol(config.collector_protocol_json_path)
    collector.save_human_report(config.collector_report_txt_path)
    files = sorted(collector.get_collected_files())

    if config.mode == "folder":
        return files

    if config.mode == "suspects":
        suspect_names = set(config.suspects)
        return [p for p in files if p.name in suspect_names]

    raise ValueError(f"Unknown mode: {config.mode!r}")


def normalize_validation_report(report: dict | None) -> dict:
    """
    Делает validation-отчет компактным и убирает дубликаты issues.
    """
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

def build_debug_payload(
    parser: DocxParser,
    validator: DocxValidator,
    file_path: Path,
) -> dict:
    """
    Собирает полную диагностическую картину по одному документу.
    Исключения не пробрасывает наружу - превращает в payload с status=error.
    """
    payload = {
        "file_name": file_path.name,
        "file_path": str(file_path),
        "error": None,

        "raw_doc_struct": None,
        "clean_struct": None,
        "raw_markers": None,
        "full_map": None,
        "primary_split": None,
        "fallback": None,
        "final_blocks": None,
        "validation": None,
    }

    try:
        raw_doc_struct = parser.read(file_path)
        payload["raw_doc_struct"] = raw_doc_struct

        if raw_doc_struct is None:
            payload["error"] = "read_failed"
            payload["validation"] = {
                "status": "error",
                "issues": ["read_failed"],
                "optional_miss": [],
            }
            return payload

        paragraphs = raw_doc_struct["paragraphs"]
        clean_struct = parser.cleanText(paragraphs)
        payload["clean_struct"] = clean_struct

        clean_paragraphs = clean_struct["clean_paragraphs"]

        raw_markers = parser.find_all_marker_indexes(clean_paragraphs)
        payload["raw_markers"] = raw_markers

        full_map = parser.build_full_docmap(raw_markers)
        payload["full_map"] = full_map

        primary_split = parser.split_into_blocks(clean_paragraphs, full_map, raw_markers)
        payload["primary_split"] = primary_split

        fallback = parser.build_fallback_doc_struct(clean_paragraphs)
        payload["fallback"] = fallback

        final_blocks = parser.get_parse_data(file_path)
        payload["final_blocks"] = final_blocks

        validation = validator.validate_doc_struct(final_blocks)
        payload["validation"] = validation

        return payload

    except Exception as exc:
        payload["error"] = f"{type(exc).__name__}: {exc}"
        payload["traceback"] = traceback.format_exc()
        payload["validation"] = {
            "status": "error",
            "issues": [f"exception:{type(exc).__name__}"],
            "optional_miss": [],
        }
        return payload


# ============================================================
# DEBUG LOG WRITER
# ============================================================

def should_write_debug(payload: dict, config: RunnerConfig) -> bool:
    """
    Решает, писать ли документ в debug log.
    """
    if not config.log_only_bad:
        return True

    validation = normalize_validation_report(payload.get("validation"))
    status = validation.get("status")
    return status in {"partial", "invalid", "ignored", "error"}


def write_debug_entry(log_path: Path, payload: dict) -> None:
    """
    Пишет один подробный debug-entry в текстовый лог.
    """
    append_text_line(log_path, "=" * 120)
    append_text_line(log_path, payload["file_name"])
    append_text_line(log_path, "=" * 120)
    append_text_line(log_path, f"FILE PATH: {payload['file_path']}")
    append_text_line(log_path)

    if payload["error"]:
        append_text_line(log_path, f"ERROR: {payload['error']}")
        append_text_line(log_path)

    if payload.get("traceback"):
        append_text_line(log_path, "TRACEBACK:")
        append_text_line(log_path, payload["traceback"])
        append_text_line(log_path)    

    append_json_block(log_path, "RAW DOC STRUCT", payload["raw_doc_struct"])
    append_json_block(log_path, "CLEAN STRUCT", payload["clean_struct"])
    append_json_block(log_path, "RAW MARKERS", payload["raw_markers"])
    append_json_block(log_path, "FULL MAP", payload["full_map"])
    append_json_block(log_path, "PRIMARY SPLIT", payload["primary_split"])
    append_json_block(log_path, "FALLBACK", payload["fallback"])
    append_json_block(log_path, "FINAL BLOCKS", payload["final_blocks"])
    append_json_block(log_path, "VALIDATION", payload["validation"])


# ============================================================
# REPORT BUILDERS
# ============================================================

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


def split_report_by_status(report_dict: dict[str, dict]) -> dict[str, dict[str, dict]]:
    """
    Делит общий report по статусам.
    """
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


def build_summary(
    report_dict: dict[str, dict],
    config: RunnerConfig,
    files: list[Path],
) -> dict:
    by_status = {
        "ok": 0,
        "partial": 0,
        "invalid": 0,
        "ignored": 0,
        "error": 0,
        "unknown": 0,
    }

    for report in report_dict.values():
        status = report.get("status", "unknown")
        if status not in by_status:
            status = "unknown"
        by_status[status] += 1

    collector_protocol = {}
    if config.collector_protocol_json_path.exists():
        with open(config.collector_protocol_json_path, "r", encoding="utf-8") as f:
            collector_protocol = json.load(f)

    return {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "mode": config.mode,
        "log_only_bad": config.log_only_bad,
        "total_files_after_collection": len(files),
        "by_status": by_status,
        "collector": {
            "accepted_packages": len(collector_protocol.get("accepted_packages", [])),
            "quarantined_packages": len(collector_protocol.get("quarantined_packages", [])),
            "copied_docx": len(collector_protocol.get("copied_docx", [])),
            "skipped_duplicates": len(collector_protocol.get("skipped_duplicates", [])),
        },
        "suspects": list(config.suspects) if config.mode == "suspects" else [],
        "collector_reports": {
            "json": str(config.collector_protocol_json_path),
            "txt": str(config.collector_report_txt_path),
        },
    }

# ============================================================
# ANALYSIS / TRIAGE HELPERS
# ============================================================

def clear_directory_files(path: Path):
    path.mkdir(parents=True, exist_ok=True)
    for item in path.iterdir():
        if item.is_file():
            item.unlink()

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


def load_plenary_titles(path: Path) -> list[str]:
    if not path.exists():
        return []

    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if isinstance(data, list):
        return [str(x) for x in data]

    return []

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

    # типовой остаточный баг: title/authors не поймались, но abstract жив
    return (
        "missing_title_ru" in issues
        and "missing_authors_ru_block" in issues
        and "missing_abstract_ru" not in issues
    )


def looks_like_missing_abstract_candidate(report: dict) -> bool:
    issues = set(report.get("issues", []))
    return "missing_abstract_ru" in issues and "missing_title_ru" not in issues


def extract_duplicate_stem(file_name: str) -> str:
    """
    Очень грубая нормализация имени файла для поиска дублей/версий.
    """
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

    # оставляем только подозрительные группы > 1 файла
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


def build_suspects_protocol(results: dict[str, dict], suspect_names: list[str]) -> dict[str, dict]:
    protocol = {}

    for file_name in suspect_names:
        payload = results.get(file_name)
        if payload is None:
            protocol[file_name] = {"error": "file_not_in_results"}
            continue

        protocol[file_name] = {
            "validation": normalize_validation_report(payload.get("validation")),
            "raw_markers": payload.get("raw_markers"),
            "full_map": payload.get("full_map"),
            "primary_split": payload.get("primary_split"),
            "fallback": payload.get("fallback"),
            "final_blocks": payload.get("final_blocks"),
            "error": payload.get("error"),
        }

    return protocol


# ============================================================
# EXTENSION HOOKS
# ============================================================

def on_before_batch(config: RunnerConfig, files: list[Path]) -> None:
    """
    Хук перед стартом пачки.
    """
    clear_text_file(config.log_path)
    append_text_line(config.log_path, "DOCX PARSER DEBUG RUN")
    append_text_line(config.log_path, f"STARTED AT: {datetime.now().isoformat(timespec='seconds')}")
    append_text_line(config.log_path, f"MODE: {config.mode}")
    append_text_line(config.log_path, f"LOG ONLY BAD: {config.log_only_bad}")
    append_text_line(config.log_path, f"FILES COUNT: {len(files)}")
    append_text_line(config.log_path)
    clear_directory_files(config.for_merge_dir)
    clear_directory_files(config.manual_review_dir)


def on_after_file(payload: dict) -> None:
    """
    Хук после обработки одного файла.
    """
    return


def on_after_batch(
    config: RunnerConfig,
    report_dict: dict[str, dict],
    summary: dict,
    results: dict[str, dict],
) -> None:
    """
    Хук после обработки всей пачки.
    """
    save_json(config.report_json_path, report_dict)
    save_json(config.summary_json_path, summary)

    buckets = split_report_by_status(report_dict)
    save_json(config.ok_json_path, buckets["ok"])
    save_json(config.partial_json_path, buckets["partial"])
    save_json(config.invalid_json_path, buckets["invalid"])
    save_json(config.ignored_json_path, buckets["ignored"])

    problem_buckets = split_problem_buckets(report_dict)
    save_json(config.applications_json_path, problem_buckets["applications"])
    save_json(config.english_only_json_path, problem_buckets["english_only"])
    save_json(config.header_bug_json_path, problem_buckets["header_bug_candidates"])
    save_json(config.missing_abstract_json_path, problem_buckets["missing_abstract_candidates"])

    duplicate_groups = build_duplicate_groups(report_dict)
    save_json(config.likely_duplicates_json_path, duplicate_groups)

    suspects_protocol = build_suspects_protocol(results, list(config.suspects))
    save_json(config.suspects_protocol_json_path, suspects_protocol)
    runtime_errors = {
        file_name: {
            "error": results[file_name].get("error"),
            "traceback": results[file_name].get("traceback"),
        }
        for file_name, report in report_dict.items()
        if report.get("status") == "error"
    }

    save_json(
        config.report_json_path.parent / "runtime_errors_detail.json",
        runtime_errors,
    )

    selector = MergeSelector(
        collected_dir=config.collected_dir,
        report_json_path=config.report_json_path,
        for_merge_dir=config.for_merge_dir,
        manual_review_dir=config.manual_review_dir,
    )
    selector.run()
    selector.save_protocol(config.merge_protocol_json_path)
    selector.save_human_report(config.merge_report_txt_path)

    restorer = StructureRestorer(
        parser=DocxParser(),
        for_merge_dir=config.for_merge_dir,
    )
    restored_blocks = restorer.run()
    restorer.save_blocks(config.restored_blocks_json_path, restored_blocks)
    restorer.save_protocol(config.restorer_protocol_json_path)
    restorer.save_human_report(config.restorer_report_txt_path)
    plenary_titles = load_plenary_titles(config.plenary_titles_json_path)

    merger = Merger(
        restored_blocks_json_path=config.restored_blocks_json_path,
        output_docx_path=config.merged_output_docx_path,
        plenary_titles=plenary_titles,
    )
    merger.run()
    
    formatter = CatalogFormatter(
        input_docx_path=config.merged_output_docx_path,
        output_docx_path=config.formatted_output_docx_path,
    )
    formatter.run()


# ============================================================
# TEST
# ============================================================

def test_docx_parser_batch():
    """
    Главный pytest-тест пакетной диагностики.
    """
    parser = DocxParser()
    validator = DocxValidator()

    files = collect_target_files(CONFIG)
    assert files, "Не найдено ни одного .docx для прогона"

    on_before_batch(CONFIG, files)

    results: dict[str, dict] = {}

    for file_path in files:
        payload = build_debug_payload(parser, validator, file_path)
        results[file_path.name] = payload

        if should_write_debug(payload, CONFIG):
            write_debug_entry(CONFIG.log_path, payload)

        on_after_file(payload)

    report_dict = build_report_dict(results)
    summary = build_summary(report_dict, CONFIG, files)

    on_after_batch(CONFIG, report_dict, summary, results)

    runtime_error_files = [
        file_name
        for file_name, report in report_dict.items()
        if report.get("status") == "error"
    ]

    # Не валим прогон сразу — сначала сохраняем отчёты и даём себе возможность их посмотреть
    if runtime_error_files:
        save_json(
            CONFIG.report_json_path.parent / "runtime_error_files.json",
            {"runtime_error_files": runtime_error_files},
        )

    # Временно отключаем падение теста
    assert not runtime_error_files