# app/core/merge_selector.py
import json
import shutil
from pathlib import Path


class MergeSelector:
    """
    Селектор кандидатов на автосклейку.

    Правила:
    - берём только файлы с минимально полной структурой:
      title_ru + authors_ru_block + abstract_ru + keywords_ru
    - имя "аннотация" НЕ является поводом для exclude
    - service-like документы исключаем
    - похожие файлы группируем по duplicate_key
    - из группы выбираем лучший по merge_score
    """

    HARD_EXCLUDE_NAME_MARKERS = (
        "заявка",
        "summary",
        "template",
        "шаблон",
        "эксперт",
        "экспертное",
        "закл",
        "заключ",
    )

    def __init__(
        self,
        collected_dir: Path,
        report_json_path: Path,
        for_merge_dir: Path,
        manual_review_dir: Path | None = None,
    ):
        self.collected_dir = collected_dir
        self.report_json_path = report_json_path
        self.for_merge_dir = for_merge_dir
        self.manual_review_dir = manual_review_dir

        self.for_merge_dir.mkdir(parents=True, exist_ok=True)
        if self.manual_review_dir is not None:
            self.manual_review_dir.mkdir(parents=True, exist_ok=True)

        self.protocol = {
            "selected_for_merge": [],
            "manual_review": [],
            "excluded": [],
        }

    # ---------------------------------------------------------
    # public
    # ---------------------------------------------------------

    def run(self):
        report = self._load_report()
        collected_files = {p.name: p for p in self.collected_dir.rglob("*.docx")}

        groups = self._build_groups(report)

        for duplicate_key, items in groups.items():
            eligible = []
            manual = []
            excluded = []

            for file_name, info in items:
                file_path = collected_files.get(file_name)

                if file_path is None:
                    excluded.append((file_name, info, "not_found_in_collected_dir"))
                    continue

                decision, reason = self._decide(file_name, info)

                if decision == "for_merge":
                    eligible.append((file_name, info, file_path, reason))
                elif decision == "manual_review":
                    manual.append((file_name, info, file_path, reason))
                else:
                    excluded.append((file_name, info, reason))

            # если в группе несколько хороших — выбираем одного лучшего
            if eligible:
                best = max(eligible, key=lambda x: x[1].get("merge_score", 0))
                best_name, best_info, best_path, best_reason = best

                target = self._copy_to_dir(best_path, self.for_merge_dir)
                self.protocol["selected_for_merge"].append({
                    "file_name": best_name,
                    "source": str(best_path),
                    "target": str(target),
                    "reason": best_reason if len(eligible) == 1 else "best_in_duplicate_group",
                    "duplicate_key": duplicate_key,
                    "merge_score": best_info.get("merge_score", 0),
                })

                for file_name, info, file_path, _ in eligible:
                    if file_name == best_name:
                        continue

                    self._send_to_manual_review(
                        file_name=file_name,
                        file_path=file_path,
                        reason="duplicate_lower_score",
                        duplicate_key=duplicate_key,
                        merge_score=info.get("merge_score", 0),
                    )

            # manual-review кандидаты
            for file_name, info, file_path, reason in manual:
                self._send_to_manual_review(
                    file_name=file_name,
                    file_path=file_path,
                    reason=reason,
                    duplicate_key=duplicate_key,
                    merge_score=info.get("merge_score", 0),
                )

            # excluded
            for file_name, info, reason in excluded:
                src = collected_files.get(file_name)
                self.protocol["excluded"].append({
                    "file_name": file_name,
                    "source": str(src) if src else None,
                    "reason": reason,
                    "duplicate_key": duplicate_key,
                    "merge_score": info.get("merge_score", 0),
                })

    def save_protocol(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.protocol, f, indent=2, ensure_ascii=False)

    def save_human_report(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)

        selected = self.protocol["selected_for_merge"]
        manual = self.protocol["manual_review"]
        excluded = self.protocol["excluded"]

        lines = []
        lines.append("MERGE SELECTOR REPORT")
        lines.append("=" * 80)
        lines.append(f"Collected dir: {self.collected_dir}")
        lines.append(f"For merge dir: {self.for_merge_dir}")
        if self.manual_review_dir is not None:
            lines.append(f"Manual review dir: {self.manual_review_dir}")
        lines.append("")

        lines.append("SUMMARY")
        lines.append("-" * 80)
        lines.append(f"Selected for merge: {len(selected)}")
        lines.append(f"Manual review:      {len(manual)}")
        lines.append(f"Excluded:           {len(excluded)}")
        lines.append("")

        lines.append("SELECTED FOR MERGE")
        lines.append("-" * 80)
        if selected:
            for idx, item in enumerate(selected, start=1):
                lines.append(f"{idx}. {item['file_name']}")
                lines.append(f"   Reason: {item['reason']}")
                lines.append(f"   Score: {item.get('merge_score')}")
                lines.append(f"   Group: {item.get('duplicate_key')}")
                lines.append("")
        else:
            lines.append("No files selected.")
            lines.append("")

        lines.append("MANUAL REVIEW")
        lines.append("-" * 80)
        if manual:
            for idx, item in enumerate(manual, start=1):
                lines.append(f"{idx}. {item['file_name']}")
                lines.append(f"   Reason: {item['reason']}")
                lines.append(f"   Score: {item.get('merge_score')}")
                lines.append(f"   Group: {item.get('duplicate_key')}")
                lines.append("")
        else:
            lines.append("No files for manual review.")
            lines.append("")

        lines.append("EXCLUDED")
        lines.append("-" * 80)
        if excluded:
            for idx, item in enumerate(excluded, start=1):
                lines.append(f"{idx}. {item['file_name']}")
                lines.append(f"   Reason: {item['reason']}")
                lines.append(f"   Score: {item.get('merge_score')}")
                lines.append(f"   Group: {item.get('duplicate_key')}")
                lines.append("")
        else:
            lines.append("No excluded files.")
            lines.append("")

        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))

    # ---------------------------------------------------------
    # internal
    # ---------------------------------------------------------

    def _load_report(self) -> dict:
        with open(self.report_json_path, "r", encoding="utf-8") as f:
            return json.load(f)

    def _build_groups(self, report: dict) -> dict[str, list[tuple[str, dict]]]:
        groups = {}

        for file_name, info in report.items():
            key = info.get("duplicate_key") or file_name.lower()
            groups.setdefault(key, []).append((file_name, info))

        return groups

    def _decide(self, file_name: str, info: dict) -> tuple[str, str]:
        status = info.get("status")
        merge_ready_minimal = info.get("merge_ready_minimal", False)
        merge_score = info.get("merge_score", 0)
        name_low = file_name.lower()

        if any(marker in name_low for marker in self.HARD_EXCLUDE_NAME_MARKERS):
            return "exclude", "excluded_by_hard_filename_marker"

        if status == "ignored":
            return "exclude", "ignored_by_parser"

        if status == "ok" and merge_ready_minimal:
            return "for_merge", "ok_and_minimal_structure_ready"

        if status == "ok" and not merge_ready_minimal:
            return "manual_review", "ok_but_minimal_structure_missing"

        if status == "partial":
            if merge_score >= 15:
                return "manual_review", "partial_but_structurally_promising"
            return "exclude", "partial_not_good_enough"

        if status == "invalid":
            return "exclude", "invalid_not_allowed_for_auto_merge"

        return "exclude", "unknown_status"

    def _send_to_manual_review(
        self,
        file_name: str,
        file_path: Path,
        reason: str,
        duplicate_key: str,
        merge_score: int,
    ):
        if self.manual_review_dir is not None:
            target = self._copy_to_dir(file_path, self.manual_review_dir)
            target_str = str(target)
        else:
            target_str = None

        self.protocol["manual_review"].append({
            "file_name": file_name,
            "source": str(file_path),
            "target": target_str,
            "reason": reason,
            "duplicate_key": duplicate_key,
            "merge_score": merge_score,
        })

    def _copy_to_dir(self, file_path: Path, target_dir: Path) -> Path:
        target = target_dir / file_path.name
        if not target.exists():
            shutil.copy(file_path, target)
            return target

        counter = 1
        while True:
            alt = target_dir / f"{file_path.stem}_{counter}{file_path.suffix}"
            if not alt.exists():
                shutil.copy(file_path, alt)
                return alt
            counter += 1