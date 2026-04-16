# app/core/structure_restorer.py
import json
from pathlib import Path
import re

class StructureRestorer:
    """
    Восстанавливает унифицированную структуру документов для будущей склейки.
    Работает только по уже отобранным кандидатам из for_merge.
    """

    def __init__(self, parser, for_merge_dir: Path):
        self.parser = parser
        self.for_merge_dir = for_merge_dir
        self.protocol = {
            "restored": [],
            "excluded": [],
        }

    # ---------------------------------------------------------
    # public
    # ---------------------------------------------------------

    def run(self) -> dict:
        result = {}

        files = sorted(self.for_merge_dir.rglob("*.docx"))
        for file_path in files:
            try:
                raw_blocks = self.parser.get_parse_data(file_path)
                restored = self._restore_document(raw_blocks, file_path)

                if restored is None:
                    self.protocol["excluded"].append({
                        "file_name": file_path.name,
                        "reason": "restore_failed_or_incomplete",
                    })
                    continue

                result[file_path.name] = restored
                self.protocol["restored"].append({
                    "file_name": file_path.name,
                    "title_ru": restored.get("title_ru"),
                })

            except Exception as e:
                self.protocol["excluded"].append({
                    "file_name": file_path.name,
                    "reason": f"exception: {type(e).__name__}: {e}",
                })

        return result

    def save_blocks(self, path: Path, blocks: dict):
        path.parent.mkdir(parents=True, exist_ok=True)

        sorted_items = sorted(
            blocks.items(),
            key=lambda kv: (kv[1].get("title_ru") or "").lower()
        )

        sorted_blocks = {k: v for k, v in sorted_items}

        with open(path, "w", encoding="utf-8") as f:
            json.dump(sorted_blocks, f, indent=2, ensure_ascii=False)

    def save_protocol(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.protocol, f, indent=2, ensure_ascii=False)

    def save_human_report(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)

        lines = []
        lines.append("STRUCTURE RESTORER REPORT")
        lines.append("=" * 80)
        lines.append(f"FOR MERGE DIR: {self.for_merge_dir}")
        lines.append("")

        restored = self.protocol["restored"]
        excluded = self.protocol["excluded"]

        lines.append("SUMMARY")
        lines.append("-" * 80)
        lines.append(f"Restored: {len(restored)}")
        lines.append(f"Excluded: {len(excluded)}")
        lines.append("")

        lines.append("RESTORED DOCUMENTS")
        lines.append("-" * 80)
        if restored:
            for idx, item in enumerate(restored, start=1):
                lines.append(f"{idx}. {item['file_name']}")
                lines.append(f"   Title: {item.get('title_ru')}")
                lines.append("")
        else:
            lines.append("No restored documents.")
            lines.append("")

        lines.append("EXCLUDED")
        lines.append("-" * 80)
        if excluded:
            for idx, item in enumerate(excluded, start=1):
                lines.append(f"{idx}. {item['file_name']}")
                lines.append(f"   Reason: {item['reason']}")
                lines.append("")
        else:
            lines.append("No excluded documents.")
            lines.append("")

        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))

    # ---------------------------------------------------------
    # internal
    # ---------------------------------------------------------

    def _restore_document(self, blocks: dict, file_path: Path) -> dict | None:
        """
        Строит нормализованную структуру документа для склейки.
        """
        title_ru = self._join_block(blocks.get("title_ru"))
        authors_ru = self._clean_block(blocks.get("authors_ru_block"))
        abstract_ru = self._join_block(blocks.get("abstract_ru"))
        keywords_ru = self._join_keywords(blocks.get("keywords_ru"))

        # Минимальный контракт для склейки
        if not title_ru or not authors_ru or not abstract_ru or not keywords_ru:
            return None

        restored = {
            "source_file": file_path.name,
            "udk": self._join_block(blocks.get("udk")),
            "title_ru": title_ru,
            "authors_ru_block": authors_ru,
            "abstract_ru": abstract_ru,
            "keywords_ru": keywords_ru,

            "title_en": self._join_block(blocks.get("title_en")),
            "authors_en_block": self._clean_block(blocks.get("authors_en_block")),
            "abstract_en": self._join_block(blocks.get("abstract_en")),
            "keywords_en": self._join_keywords(blocks.get("keywords_en")),

            "reference_block": self._normalize_reference_block(blocks.get("reference_block")),
        }

        return restored

    def _normalize_reference_block(self, ref_block):
        cleaned = self._deduplicate_references(ref_block)
        if not cleaned:
            return None

        bad_headers = {
            "список литературы",
            "список источников",
            "references",
            "reference",
            "литература",
        }

        result = []
        for line in cleaned:
            text = " ".join(str(line).replace("\xa0", " ").split()).strip()
            if not text:
                continue

            low = text.lower()

            if low in bad_headers:
                continue

            if "разрыв страницы" in low:
                continue

            if not self._looks_like_reference_line(text):
                continue

            text = self._strip_reference_number(text)
            result.append(text)

        if not result:
            return None

        # Перенумерация в единый вид
        result = [f"[{i}] {line}" for i, line in enumerate(result, start=1)]
        return result

    def _looks_like_reference_line(self, text: str) -> bool:
        low = text.lower()

        if re.match(r"^\[\d+\]", text):
            return True
        if re.match(r"^\d+[\.\)]\s+", text):
            return True
        if "doi" in low:
            return True
        if "//" in text:
            return True
        if "http://" in low or "https://" in low:
            return True
        if re.search(r"\bvol\.?\b", low):
            return True
        if re.search(r"\bno\.?\b", low):
            return True
        if re.search(r"\bpp?\.?\s*\d+", low):
            return True
        if re.search(r"№\s*\d+", text):
            return True
        if re.search(r"\bdoi\.org\b", low):
            return True

        # мягкий фолбэк: похоже на библиографию, если длинная строка с авторами/изданием
        if len(text) > 40 and any(mark in low for mark in ("м.", "moscow", "journal", "изд", "press", "conference")):
            return True

        return False

    def _strip_reference_number(self, text: str) -> str:
        text = re.sub(r"^\[\d+\]\s*", "", text)   # [1] ...
        text = re.sub(r"^\d+\.\s*", "", text)     # 1. ...
        text = re.sub(r"^\d+\)\s*", "", text)     # 1) ...
        return text.strip()

    def _deduplicate_references(self, ref_block):
        cleaned = self._clean_block(ref_block)
        if not cleaned:
            return None

        uniq = []
        seen = set()

        for line in cleaned:
            key = " ".join(line.lower().split())
            if key in seen:
                continue
            seen.add(key)
            uniq.append(line)

        return uniq if uniq else None

    def _clean_block(self, block):
        if not block:
            return None

        cleaned = []
        bad_exact = {
            "annotation",
            "abstract",
            "keywords",
            "аннотация",
            "ключевые слова",
            "references",
            "список литературы",
        }

        for line in block:
            if line is None:
                continue

            text = " ".join(str(line).replace("\xa0", " ").split()).strip()
            if not text:
                continue

            low = text.lower()

            if low in bad_exact:
                continue
            if "разрыв страницы" in low:
                continue
            if low.startswith("the study was conducted"):
                continue

            cleaned.append(text)

        return cleaned if cleaned else None

    def _join_block(self, block):
        cleaned = self._clean_block(block)
        if not cleaned:
            return None
        return " ".join(cleaned)

    def _join_keywords(self, block):
        cleaned = self._clean_block(block)
        if not cleaned:
            return None

        text = " ".join(cleaned).strip()

        prefixes = (
            "ключевые слова:",
            "ключевые слова.",
            "keywords:",
            "keywords.",
        )

        low = text.lower()
        for prefix in prefixes:
            if low.startswith(prefix):
                text = text[len(prefix):].strip()
                break

        # нормализуем запятые
        text = text.replace(",", ", ")
        text = " ".join(text.split())

        return text        
