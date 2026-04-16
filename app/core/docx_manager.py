# docx_manager.py
from docx import Document
from pathlib import Path
import re
import json
import deep_translator as translator

CITYIES_PATH = r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\config\russian-cities.json"
token_path = r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\config\token_txt.txt"
with open(token_path, "r") as f:
    API_KEY = f.readline()

PARA_KIND = {
    "udk",
    "abstract_ru_marker",
    "keywords_ru_marker",
    "abstract_en_marker",
    "keywords_en_marker",
    "reference_marker",
    "reference_item",
    "author_ru_line",
    "author_en_line",
    "org_line",
    "email_line",
    "service_line",
    "unknown",
}

class DocxParser:
    def __init__(self):
        self.load_cityies_set()
        

    def read(self, path: Path) -> dict:  #{paragraph_list: list[str], raw_text: str}
        try:
            doc = Document(path)
        except Exception as e:
            print(f"Не удалось прочитать {path.name}: {e}")
            return None    
        paragraphs = doc.paragraphs
        paragraph_list = []
        for paragraph in paragraphs:
            text = paragraph.text.strip()
            if text:
                paragraph_list.append(text)
        raw_text = "\n".join(paragraph_list)
        return  {
                    "raw_text": raw_text,
                    "paragraphs": paragraph_list
                }

    def cleanText(self, paragraphs: list[str]):
        clean_paragraphs = []
        for paragraph in paragraphs:
            text = paragraph.replace("\xa0", " ")
            text = " ".join(text.split())
            if not text:
                continue
            if len(text) > 2 and len(set(text)) == 1 and text[0] in "-_*":
                continue
            clean_paragraphs.append(text)
        clean_text = "\n".join(clean_paragraphs)
        return  {
                    "clean_text": clean_text,
                    "clean_paragraphs": clean_paragraphs
                }

    def extract_email(self, clean_text: str) -> str | None:
        pattern = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]+"
        match = re.search(pattern, clean_text)
        if match:
            return match.group(0)
        else:
            return None       
    
    def extract_block(self, clean_paragraphs: list[str], start_par = 0, finish_par = None) -> list[str] | None:
        return clean_paragraphs[start_par:finish_par]
  

    def normalize_text(self, text: str) -> str:
        return " ".join(text.replace("\xa0", " ").split()).strip().lower()
    
    def is_garbage_title_en(self, text: str):
        bad = (
            "doi", "//", "vol", "pp", "journal",
            "transactions", "conference",
            "список", "литература"
        )
        text_low = text.lower()
        return any(b in text_low for b in bad)

    def is_garbage_title_line(self, text: str) -> bool:
        text = " ".join(text.replace("\xa0", " ").split()).strip()
        if not text:
            return True

        if text in {".", ",", "-", "–", "_", ":", ";"}:
            return True

        # почти одна пунктуация
        alpha_count = sum(ch.isalpha() for ch in text)
        if alpha_count == 0:
            return True

        text_norm = text.lower()

        bad_starts = (
            "аннотация",
            "annotation",
            "abstract",
            "ключевые слова",
            "keywords",
            "список литературы",
            "список источников",
            "литература",
            "references",
            "исследование выполнено",
            "the study was conducted",
            "работа выполнена",
        )
        if any(text_norm.startswith(x) for x in bad_starts):
            return True

        if self.is_email_line(text):
            return True

        return False

    def is_email_line(self, text: str) -> bool:
        return self.extract_email(text) is not None

    def is_spin_line(self, text: str) -> bool:
        text = " ".join(text.replace("\xa0", " ").split()).strip()
        if not text:
            return False
        return text.upper().startswith("SPIN")

    def is_author_ru_line(self, text: str) -> bool:
        text = " ".join(text.replace("\xa0", " ").split()).strip()
        if not text:
            return False

        # не считаем служебные маркеры авторами
        text_norm = text.lower()
        bad_starts = (
            "удк",
            "аннотация",
            "ключевые слова",
            "список литературы",
            "список источников",
            "литература",
            "abstract",
            "keywords",
        )
        if any(text_norm.startswith(x) for x in bad_starts):
            return False

        # Фамилия И.О.
        if re.search(r"\b[А-ЯЁ][а-яё-]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.\s*\d*(?:\(\*\))?\*?\b", text):
            return True

        # И.О. Фамилия
        if re.search(r"\b[А-ЯЁ]\.\s*[А-ЯЁ]\.\s*[А-ЯЁа-яё-]+\s*\d*(?:\(\*\))?\*?\b", text):
            return True

        # Полное ФИО: Имя Отчество Фамилия / Фамилия Имя Отчество
        if re.search(r"\b[А-ЯЁ][а-яё-]+\s+[А-ЯЁ][а-яё-]+\s+[А-ЯЁ][а-яё-]+\d*(?:\(\*\))?\*?\b", text):
            return True

        # Несколько авторов через запятую с инициалами
        if re.search(r"[А-ЯЁ][а-яё-]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.", text) and "," in text:
            return True
        if re.search(r"[А-ЯЁ]\.\s*[А-ЯЁ]\.\s*[А-ЯЁа-яё-]+", text) and "," in text:
            return True

        # строки вроде "Овсянников В.М."
        if re.fullmatch(r"[А-ЯЁ][а-яё-]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.?", text):
            return True

        # строки с цифрой аффилиации: "Аитов Василий Григорьевич3"
        if re.fullmatch(r"[А-ЯЁ][а-яё-]+\s+[А-ЯЁ][а-яё-]+\s+[А-ЯЁ][а-яё-]+\d+", text):
            return True

        # строки с автором и email
        if self.extract_email(text) and re.search(r"[А-ЯЁA-Z]", text):
            if re.search(r"\b[А-ЯЁ][а-яё-]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.", text):
                return True
            if re.search(r"\b[А-ЯЁ][а-яё-]+\s+[А-ЯЁ][а-яё-]+\s+[А-ЯЁ][а-яё-]+", text):
                return True

        return False

    def looks_like_author_fallback_line(self, text: str) -> bool:
        text = " ".join(text.replace("\xa0", " ").split()).strip()
        if not text:
            return False

        text_norm = text.lower()

        bad_starts = (
            "аннотация",
            "annotation",
            "abstract",
            "ключевые слова",
            "keywords",
            "список литературы",
            "список источников",
            "литература",
            "references",
            "исследование выполнено",
            "the study was conducted",
            "работа выполнена",
        )
        if any(text_norm.startswith(x) for x in bad_starts):
            return False

        # одиночный автор: Овсянников В.М.
        if re.fullmatch(r"[А-ЯЁ][а-яё-]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.?", text):
            return True

        # ФИО/инициалы с индексами и звёздочками
        if re.search(r"[А-ЯЁA-Z][A-Za-zА-Яа-яЁё.-]+\s+[А-ЯЁA-Z]\.\s*[А-ЯЁA-Z]\.?", text):
            return True

        # англ/рус строка с несколькими авторами через запятую
        if "," in text and any(ch.isalpha() for ch in text):
            if self.extract_email(text) is None:
                return True

        # author + email на одной строке
        if self.extract_email(text):
            if re.search(r"[А-ЯЁA-Z][A-Za-zА-Яа-яЁё.-]+", text):
                return True

        return False


    def is_author_en_line(self, text: str) -> bool:
        text = " ".join(text.replace("\xa0", " ").split()).strip()
        if not text:
            return False

        text_norm = text.lower()
        bad_starts = (
            "abstract",
            "keywords",
            "literature",
            "references",
        )
        if any(text_norm.startswith(x) for x in bad_starts):
            return False

        # S.N. Lelyavin
        if re.search(r"\b[A-Z]\.\s*[A-Z]\.\s*[A-Z][a-z-]+\b", text):
            return True

        # Lelyavin S.N.
        if re.search(r"\b[A-Z][a-z-]+\s+[A-Z]\.\s*[A-Z]\.\b", text):
            return True

        # Full name
        if re.fullmatch(r"[A-Z][a-z-]+(?:\s+[A-Z][a-z-]+){1,3}\d*(?:\(\*\))?\*?", text):
            return True

        # Several authors
        if "," in text and re.search(r"[A-Z][a-z-]+\s+[A-Z][a-z-]+", text):
            return True
        if "," in text and re.search(r"[A-Z]\.\s*[A-Z]\.\s*[A-Z][a-z-]+", text):
            return True

        # name + email
        if self.extract_email(text) and re.search(r"[A-Z][a-z]", text):
            return True

        return False

    def is_reference_marker_line(self, text: str) -> bool:
        text = " ".join(text.replace("\xa0", " ").split()).strip()
        if not text:
            return False

        text_norm = text.lower()
        return text_norm in {
            "список источников",
            "список литературы",
            "литература",
            "references",
            "literature",
        }

    def is_reference_item_line(self, text: str) -> bool:
        text = " ".join(text.replace("\xa0", " ").split()).strip()
        if not text:
            return False

        text_norm = text.lower()

        if re.match(r"^\[\d+\]", text):
            return True
        if re.match(r"^\d+\.\s+", text):
            return True
        if "//" in text:
            return True
        if re.search(r"\bdoi\b", text_norm):
            return True
        if re.search(r"\bvol\.?\b", text_norm):
            return True
        if re.search(r"\bno\.?\b", text_norm):
            return True
        if re.search(r"\bp\.?\s*\d", text_norm):
            return True
        if re.search(r"\bстр\.?\b", text_norm):
            return True
        if re.search(r"\bвып\.?\b", text_norm):
            return True
        if re.search(r"№\s*\d+", text):
            return True

        return False

    def is_keywords_ru_line(self, text: str) -> bool:
        t = self.normalize_text(text)
        return t.startswith("ключевые слова")


    def is_keywords_en_line(self, text: str) -> bool:
        t = self.normalize_text(text)
        return t.startswith("keywords")


    def is_abstract_header_ru(self, text: str) -> bool:
        t = self.normalize_text(text)
        return t == "аннотация"


    def is_abstract_header_en(self, text: str) -> bool:
        t = self.normalize_text(text)
        return t == "abstract"


    def is_abstract_ru_line(self, text: str) -> bool:
        t = self.normalize_text(text)
        return (
            t == "аннотация"
            or t.startswith("аннотация.")
            or t.startswith("аннотация:")
            or t.startswith("аннотация ")
        )


    def is_abstract_en_line(self, text: str) -> bool:
        t = self.normalize_text(text)
        return t == "abstract" or t.startswith("abstract.") or t.startswith("abstract:") or t.startswith("annotation.") or t.startswith("annotation:") or t.startswith("annotation ") or t.startswith("abstract ")

    
   
    
    

    def find_all_marker_indexes(self, clean_paragraphs: list[str]) -> dict[str, int | None]:
        doc_map = {
            "udk": None,
            "abstract_ru": None,
            "keywords_ru": None,
            "abstract_en": None,
            "keywords_en": None,
            "reference_block": None,
        }

        for i, p in enumerate(clean_paragraphs):
            if doc_map["udk"] is None and self.normalize_text(p).startswith("удк"):
                doc_map["udk"] = i
                continue

            if doc_map["abstract_ru"] is None and self.is_abstract_ru_line(p):
                doc_map["abstract_ru"] = i
                continue

            if doc_map["keywords_ru"] is None and self.is_keywords_ru_line(p):
                doc_map["keywords_ru"] = i
                continue

            if doc_map["abstract_en"] is None and self.is_abstract_en_line(p):
                doc_map["abstract_en"] = i
                continue

            if doc_map["keywords_en"] is None and self.is_keywords_en_line(p):
                doc_map["keywords_en"] = i
                continue

            if doc_map["reference_block"] is None and self.is_reference_marker_line(p):
                doc_map["reference_block"] = i
                continue

        return doc_map
      

    def resolve_block_range(self, key: str, doc_map_raw: dict):
        start = doc_map_raw.get(key)
        if start is None:
            return None

        next_candidates = []

        if key == "abstract_ru":
            for nxt in ("keywords_ru", "abstract_en", "keywords_en", "reference_block", "body_block"):
                idx = doc_map_raw.get(nxt)
                if idx is not None and idx > start:
                    next_candidates.append(idx)

        elif key == "abstract_en":
            for nxt in ("keywords_en", "reference_block", "body_block"):
                idx = doc_map_raw.get(nxt)
                if idx is not None and idx > start:
                    next_candidates.append(idx)

        elif key == "body_block":
            for nxt in ("reference_block",):
                idx = doc_map_raw.get(nxt)
                if idx is not None and idx > start:
                    next_candidates.append(idx)

        elif key == "reference_block":
            return (start, None)

        end = min(next_candidates) if next_candidates else None
        return (start, end)
    
    def get_section_end(self, start_idx: int | None, candidates: list[int | None]) -> int | None:
        if start_idx is None:
            return None

        valid = [idx for idx in candidates if idx is not None and idx > start_idx]
        return min(valid) if valid else None

    
    def recover_header_from_pre_abstract_block(
        self,
        clean_paragraphs: list[str],
        doc_map_raw: dict,
        doc_struct: dict,
    ) -> dict:
        if doc_struct.get("title_ru") not in (None, []):
            return doc_struct
        if doc_struct.get("authors_ru_block") not in (None, []):
            return doc_struct

        abs_idx = doc_map_raw.get("abstract_ru")
        if abs_idx is None:
            return doc_struct

        start = 0 if doc_map_raw.get("udk") is None else doc_map_raw["udk"] + 1
        block = clean_paragraphs[start:abs_idx]
        if not block:
            return doc_struct

        # чистим мусор
        candidates = []
        for p in block:
            t = " ".join(p.replace("\xa0", " ").split()).strip()
            if not t:
                continue
            if self.is_service_header(t):
                continue
            if self.is_garbage_title_line(t):
                continue
            candidates.append(t)

        if not candidates:
            return doc_struct

        # title = первая не-author/org/email строка
        title = None
        authors = []

        for i, line in enumerate(candidates):
            kind = self.classify_author_line(line)
            if kind not in ("name", "email", "org_text", "spin_code"):
                title = [line]
                tail = candidates[i + 1:]
                for q in tail:
                    q_kind = self.classify_author_line(q)
                    if q_kind in ("name", "email", "org_text", "spin_code") or self.looks_like_author_fallback_line(q):
                        authors.append(q)
                break

        # если title не найден, берём первую строку как title, остальное как authors-like
        if title is None:
            title = [candidates[0]]
            for q in candidates[1:]:
                q_kind = self.classify_author_line(q)
                if q_kind in ("name", "email", "org_text", "spin_code") or self.looks_like_author_fallback_line(q):
                    authors.append(q)

        if title:
            doc_struct["title_ru"] = title
        if authors:
            doc_struct["authors_ru_block"] = authors if authors else None

        return doc_struct

    def extract_header_block(self, clean_paragraphs: list[str], start: int, end: int):
        if start is None or end is None or start >= end:
            return None, None

        block = clean_paragraphs[start:end]
        if not block:
            return None, None

        title = None
        authors = []

        i = 0

        # 1. пропускаем явный служебный мусор в начале
        while i < len(block):
            p = block[i]
            text_norm = self.normalize_text(p)

            if not p.strip():
                i += 1
                continue

            if self.is_service_header(p):
                i += 1
                continue

            if self.is_email_line(p):
                i += 1
                continue

            # не даем мусору/авторам стать title
            if self.classify_author_line(p) in ("name", "email", "spin_code", "org_text"):
                i += 1
                continue

            if self.is_garbage_title_line(p):
                i += 1
                continue

            title = [p]
            i += 1
            break

        if title is None:
            return None, None

        # 2. всё подряд после title, пока это похоже на author/org/email/spin
        while i < len(block):
            p = block[i]
            kind = self.classify_author_line(p)

            if kind in ("name", "email", "spin_code", "org_text"):
                authors.append(p)
                i += 1
                continue

            # разрешаем одну "непонятную", но очень author-like строку
            # например: "Lobanov M.V. 1 *, Belichenko M.V. 1"
            # или "Овсянников В.М."
            if self.looks_like_author_fallback_line(p):
                authors.append(p)
                i += 1
                continue

            break

         # 🚨 FALLBACK: если не нашли title — берём первую вменяемую строку
        if title is None:
            for p in block:
                if not self.is_garbage_title_line(p):
                    if not self.is_service_header(p):
                        title = [p]
                        break

        # 🚨 FALLBACK: если нет authors — пробуем собрать хоть что-то
        if title and not authors:
            for p in block:
                if self.looks_like_author_fallback_line(p):
                    authors.append(p)

        return title, (authors if authors else None)



    def extract_body_block_from_map(self, clean_paragraphs: list[str], doc_map_raw: dict) -> list[str] | None:
        start_candidates = []

        for key in ("keywords_en", "abstract_en", "keywords_ru", "abstract_ru"):
            idx = doc_map_raw.get(key)
            if idx is not None:
                start_candidates.append((key, idx))

        if start_candidates:
            last_key, last_idx = max(start_candidates, key=lambda x: x[1])

            if last_key in ("keywords_ru", "keywords_en"):
                start = last_idx + 1
            else:
                end = self.get_section_end(
                    last_idx,
                    [
                        doc_map_raw.get("keywords_ru"),
                        doc_map_raw.get("keywords_en"),
                        doc_map_raw.get("abstract_en"),
                        doc_map_raw.get("reference_block"),
                    ]
                )
                start = end if end is not None else last_idx + 1
        else:
            start = 0

        ref_idx = doc_map_raw.get("reference_block")
        end = ref_idx if ref_idx is not None else len(clean_paragraphs)

        if start >= end:
            return None

        body = clean_paragraphs[start:end]
        return body if body else None
    
    def cleanup_abstract_markers(self, doc_struct: dict) -> dict:
        for key in ("abstract_ru", "abstract_en"):
            block = doc_struct.get(key)
            if not block:
                continue

            cleaned = []
            for i, line in enumerate(block):
                t = " ".join(line.replace("\xa0", " ").split()).strip()

                if i == 0:
                    low = t.lower()

                    if low in {"аннотация", "annotation", "abstract", "аnnotation"}:
                        continue

                    if low.startswith("аннотация."):
                        cleaned.append(t[len("аннотация."):].strip())
                        continue
                    if low.startswith("аннотация:"):
                        cleaned.append(t[len("аннотация:"):].strip())
                        continue

                    if low.startswith("abstract."):
                        cleaned.append(t[len("abstract."):].strip())
                        continue
                    if low.startswith("abstract:"):
                        cleaned.append(t[len("abstract:"):].strip())
                        continue

                    if low.startswith("annotation."):
                        cleaned.append(t[len("annotation."):].strip())
                        continue
                    if low.startswith("annotation:"):
                        cleaned.append(t[len("annotation:"):].strip())
                        continue

                cleaned.append(t)

            cleaned = [x for x in cleaned if x]
            doc_struct[key] = cleaned if cleaned else None

        return doc_struct

    def postprocess_inline_abstracts(self, doc_struct: dict) -> dict:
        abs_ru = doc_struct.get("abstract_ru")
        if abs_ru:
            first = abs_ru[0]
            norm = self.normalize_text(first)
            if norm.startswith("аннотация.") or norm.startswith("аннотация:"):
                stripped = re.sub(r"^\s*аннотация(?:\s*[\.:]|\s+)\s*", "", first, flags=re.IGNORECASE).strip()
                doc_struct["abstract_ru"] = [stripped] + abs_ru[1:] if stripped else abs_ru[1:]

        abs_en = doc_struct.get("abstract_en")
        if abs_en:
            first = abs_en[0]
            norm = self.normalize_text(first)
            if norm.startswith("abstract.") or norm.startswith("abstract:"):
                stripped = re.sub(r"^\s*abstract(?:\s*[\.:]|\s+)\s*", "", first, flags=re.IGNORECASE).strip()
                doc_struct["abstract_en"] = [stripped] + abs_en[1:] if stripped else abs_en[1:]

        return doc_struct
    
    def recover_header_from_abstract_block(self, doc_struct: dict) -> dict:
        abs_block = doc_struct.get("abstract_ru") or []
        if len(abs_block) < 3:
            return doc_struct

        first = self.normalize_text(abs_block[0])
        if first != "аннотация":
            return doc_struct

        title = None
        authors = []
        abstract = []

        i = 1

        if i < len(abs_block):
            title = [abs_block[i]]
            i += 1

        while i < len(abs_block):
            line = abs_block[i]
            kind = self.classify_author_line(line)
            if kind in ("name", "email", "spin_code", "org_text"):
                authors.append(line)
                i += 1
                continue
            break

        abstract = abs_block[i:] if i < len(abs_block) else None

        if title:
            doc_struct["title_ru"] = title
        if authors:
            doc_struct["authors_ru_block"] = authors
        if abstract:
            doc_struct["abstract_ru"] = abstract

        return doc_struct

    def recover_abstract_from_body(self, doc_struct: dict) -> dict:
        if doc_struct.get("abstract_ru") not in (None, []):
            return doc_struct

        body = doc_struct.get("body_block") or []
        if body:
            first = " ".join(body[0].replace("\xa0", " ").split()).strip()
            low = first.lower()

            bad_starts = (
                "список литературы",
                "список источников",
                "литература",
                "references",
                "reference",
                "keywords",
                "ключевые слова",
                "введение",
                "introduction",
            )

            alpha_count = sum(ch.isalpha() for ch in first)
            if alpha_count >= 40 and not any(low.startswith(x) for x in bad_starts):
                doc_struct["abstract_ru"] = [first]
                doc_struct["body_block"] = body[1:] if len(body) > 1 else None
                return doc_struct

        # fallback-body from build_fallback_doc_struct may already hold abstract+keywords
        return doc_struct
    

    def merge_fallback_abstract_from_body(self, doc_struct: dict, fallback_struct: dict) -> dict:
        if doc_struct.get("abstract_ru") not in (None, []):
            return doc_struct

        fb_body = (fallback_struct or {}).get("body_block") or []
        if not fb_body:
            return doc_struct

        first = " ".join(fb_body[0].replace("\xa0", " ").split()).strip()
        low = first.lower()

        bad_starts = (
            "литература",
            "references",
            "keywords",
            "ключевые слова",
            "введение",
            "introduction",
        )

        alpha_count = sum(ch.isalpha() for ch in first)
        if alpha_count >= 40 and not any(low.startswith(x) for x in bad_starts):
            doc_struct["abstract_ru"] = [first]
            return doc_struct

        return doc_struct

    def recover_en_from_reference_block(self, doc_struct: dict) -> dict:
        ref_block = doc_struct.get("reference_block") or []
        if not ref_block:
            return doc_struct

        if doc_struct.get("title_en") not in (None, []):
            return doc_struct

        # ищем начало EN-секции внутри reference_block
        start_idx = None
        for i, line in enumerate(ref_block):
            t = " ".join(line.replace("\xa0", " ").split()).strip()
            low = t.lower()

            if not t:
                continue
            if self.is_reference_item_line(t):
                continue
            if low in {"references", "literature", "reference"}:
                continue
            if self.is_garbage_title_line(t):
                continue

            latin_letters = sum(ch.isascii() and ch.isalpha() for ch in t)
            cyr_letters = sum(("А" <= ch <= "я") or ch in "Ёё" for ch in t)

            # EN title-like
            if latin_letters >= 8 and cyr_letters == 0:
                start_idx = i
                break

        if start_idx is None:
            return doc_struct

        tail = ref_block[start_idx:]
        head = ref_block[:start_idx]

        if not tail:
            return doc_struct

        # title_en
        title_en = None
        authors_en = []
        abstract_en = None
        keywords_en = None

        i = 0
        if not self.is_garbage_title_line(tail[0]):
            title_en = [tail[0]]
            i = 1

        # authors_en
        while i < len(tail):
            line = tail[i]
            line_norm = self.normalize_text(line)

            if line_norm.startswith(("annotation", "abstract", "аnnotation")):
                break
            if line_norm.startswith("keywords"):
                break
            if line_norm.startswith("references"):
                break
            if line_norm.startswith("the study was conducted"):
                break

            if self.is_reference_item_line(line):
                break

            kind = self.classify_author_line(line)
            if kind in ("name", "email", "org_text", "spin_code") or self.looks_like_author_fallback_line(line):
                authors_en.append(line)
                i += 1
                continue

            # если строка длинная и похожа на абзац, это уже не authors
            alpha_count = sum(ch.isalpha() for ch in line)
            if alpha_count >= 60:
                break

            i += 1

        # abstract_en
        while i < len(tail):
            line = tail[i]
            line_norm = self.normalize_text(line)

            if line_norm.startswith("annotation") or line_norm.startswith("abstract"):
                text = line
                if ":" in text:
                    _, right = text.split(":", 1)
                    abstract_en = [right.strip()] if right.strip() else None
                elif "." in text:
                    left, right = text.split(".", 1)
                    abstract_en = [right.strip()] if right.strip() else None
                else:
                    abstract_en = tail[i+1:i+2] or None
                i += 1
                break

            i += 1

        # keywords_en
        for j in range(i, len(tail)):
            line = tail[j]
            if self.is_keywords_en_line(line):
                keywords_en = [line]
                break

        if title_en and self.is_garbage_title_en(title_en[0]):
            title_en = None

        if title_en:
            doc_struct["title_en"] = title_en
        if authors_en:
            doc_struct["authors_en_block"] = authors_en
        if abstract_en:
            doc_struct["abstract_en"] = abstract_en
        if keywords_en:
            doc_struct["keywords_en"] = keywords_en

        # reference_block режем обратно
        # убираем EN-мусор полностью
        clean_head = []
        saw_ref_header = False

        for line in head:
            t = " ".join(line.replace("\xa0", " ").split()).strip()
            low = t.lower()

            if low in {"литература", "references", "reference", "список литературы", "список источников"}:
                saw_ref_header = True
                clean_head.append(line)
                continue

            if self.is_reference_item_line(t):
                clean_head.append(line)
                continue

        # если в head остались только библиографические строки, оставляем их
        doc_struct["reference_block"] = clean_head if clean_head else None

        return doc_struct

    def build_header_ranges(self, doc_map_raw: dict, before_key, after_key):
        before_idx = doc_map_raw.get(before_key)
        after_idx = doc_map_raw.get(after_key)

        if after_idx is None:
            return None, None

        start = 0 if before_idx is None else before_idx + 1
        end = after_idx

        if start >= end:
            return None, None

        title_range = (start, start + 1)

        authors_range = None
        if start + 1 < end:
            authors_range = (start + 1, end)

        return title_range, authors_range


    def build_full_docmap(self, doc_map_raw: dict):
        full_map = {
            "udk": None,
            "title_ru": None,
            "authors_ru_block": None,
            "abstract_ru": None,
            "keywords_ru": None,
            "title_en": None,
            "authors_en_block": None,
            "abstract_en": None,
            "keywords_en": None,
            "body_block": None,
            "reference_block": None,
        }

        # одиночные строки
        for key in ("udk", "keywords_ru", "keywords_en"):
            idx = doc_map_raw.get(key)
            if idx is not None:
                full_map[key] = (idx, idx + 1)

        # обычные блоки
        for key in ("abstract_ru", "abstract_en", "reference_block"):
            full_map[key] = self.resolve_block_range(key, doc_map_raw)

        # RU header = между UDK и abstract_ru
        abs_ru_idx = doc_map_raw.get("abstract_ru")
        if abs_ru_idx is not None:
            ru_start = 0 if doc_map_raw.get("udk") is None else doc_map_raw["udk"] + 1
            full_map["title_ru"] = ("__HEADER_RU__", ru_start, abs_ru_idx)
            full_map["authors_ru_block"] = ("__HEADER_RU__", ru_start, abs_ru_idx)

        # EN header = между keywords_ru / abstract_ru и abstract_en
        abs_en_idx = doc_map_raw.get("abstract_en")
        if abs_en_idx is not None:
            if doc_map_raw.get("keywords_ru") is not None:
                en_start = doc_map_raw["keywords_ru"] + 1
            elif doc_map_raw.get("abstract_ru") is not None:
                ru_abs_end = self.resolve_block_range("abstract_ru", doc_map_raw)
                en_start = ru_abs_end[1] if ru_abs_end and ru_abs_end[1] is not None else doc_map_raw["abstract_ru"] + 1
            else:
                en_start = 0
            full_map["title_en"] = ("__HEADER_EN__", en_start, abs_en_idx)
            full_map["authors_en_block"] = ("__HEADER_EN__", en_start, abs_en_idx)

        return full_map          

  
    def split_into_blocks(self, clean_paragraphs: list[str], doc_map_full: dict, doc_map_raw: dict) -> dict:
        doc_struct = {
            "udk": None,
            "title_ru": None,
            "authors_ru_block": None,
            "abstract_ru": None,
            "keywords_ru": None,
            "title_en": None,
            "authors_en_block": None,
            "abstract_en": None,
            "keywords_en": None,
            "body_block": None,
            "reference_block": None,
        }

        for key, value in doc_map_full.items():
            if value is None:
                continue

            if isinstance(value, tuple) and value and value[0] == "__HEADER_RU__":
                _, start, end = value
                title, authors = self.extract_header_block(clean_paragraphs, start, end)
                if title and self.is_garbage_title_line(title[0]):
                    title = None
                doc_struct["title_ru"] = title
                doc_struct["authors_ru_block"] = authors
                continue

            if isinstance(value, tuple) and value and value[0] == "__HEADER_EN__":
                _, start, end = value
                title, authors = self.extract_header_block(clean_paragraphs, start, end)
                if title and self.is_garbage_title_line(title[0]):
                    title = None
                doc_struct["title_en"] = title
                doc_struct["authors_en_block"] = authors
                continue

            start, end = value
            doc_struct[key] = self.extract_block(clean_paragraphs, start, end)

        doc_struct["body_block"] = self.extract_body_block_from_map(clean_paragraphs, doc_map_raw)
        return doc_struct

    def clean_en_tail(blocks):

        ref_markers = ("references", "список", "литература")

        def is_ref_line(s):
            s = s.lower()
            return any(m in s for m in ref_markers)

        # режем abstract_en
        if blocks["abstract_en"]:
            clean = []
            for line in blocks["abstract_en"]:
                if is_ref_line(line):
                    break
                clean.append(line)
            blocks["abstract_en"] = clean or None

        # режем authors_en_block
        if blocks["authors_en_block"]:
            clean = []
            for line in blocks["authors_en_block"]:
                if is_ref_line(line):
                    break
                clean.append(line)
            blocks["authors_en_block"] = clean or None

        return blocks

    def is_body_block(self, text: str) -> bool:
        text = " ".join(text.replace("\xa0", " ").split()).strip()
        if not text:
            return False

        text_norm = text.lower()

        bad_cases = (
            "ключевые слова",
            "список источников",
            "список литературы",
            "cписок литературы",
            "литература",
            "аннотация",
            "abstract",
            "keywords",
        )
        if any(bc in text_norm for bc in bad_cases):
            return False

        if self.classify_author_line(text) in ("name", "spin_code", "org_text"):
            return False

        return len(text) >= 80 and text.count(" ") >= 8


    def is_reference_block(self, text: str) -> bool:
        if text.strip().lower().startswith("references"):
            return True
        return self.is_reference_marker_line(text) or self.is_reference_item_line(text)


    def deduplicate_blocks(blocks):

        if blocks["abstract_en"] and blocks["reference_block"]:
            abs_text = " ".join(blocks["abstract_en"])

            cleaned_ref = []
            for line in blocks["reference_block"]:
                if line not in abs_text:
                    cleaned_ref.append(line)

            blocks["reference_block"] = cleaned_ref

        return blocks

    def is_service_header(self, text: str) -> bool:
        text_norm = self.normalize_text(text)

        bad_starts = (
            "аннотация",
            "тезисы",
            "тезисы доклада",
            "тезисы к докладу",
            "заявка",
            "заключение",
            "экспертное заключение",
            "application",
            "summary",
            "abstract",
        )

        return len(text_norm) < 40 and any(text_norm.startswith(item) for item in bad_starts)


    def build_fallback_doc_struct(self, clean_paragraphs: list[str]) -> dict:
        doc_struct = {
            "udk": None,
            "title_ru": None,
            "authors_ru_block": None,
            "abstract_ru": None,
            "keywords_ru": None,
            "title_en": None,
            "authors_en_block": None,
            "abstract_en": None,
            "keywords_en": None,
            "body_block": None,
            "reference_block": None,
        }

        if not clean_paragraphs:
            return doc_struct

        n = len(clean_paragraphs)

        # --- 0. UDK
        start_idx = 0
        if self.normalize_text(clean_paragraphs[0]).startswith("удк"):
            doc_struct["udk"] = [clean_paragraphs[0]]
            start_idx = 1

        # --- 1. title_ru = всё подряд до первого author/org/email/keywords/reference/english
        i = start_idx
        title_lines = []

        while i < n:
            p = clean_paragraphs[i]

            if self.is_keywords_ru_line(p) or self.is_keywords_en_line(p):
                break
            if self.is_reference_block(p):
                break
            if self.is_abstract_ru_line(p) or self.is_abstract_en_line(p):
                break

            if self.classify_author_line(p) in ("name", "spin_code", "org_text"):
                break
            if "@" in p:
                break
            if self.looks_like_english_title(p):
                break

            title_lines.append(p)
            i += 1

        if not title_lines:
            return doc_struct

        doc_struct["title_ru"] = title_lines


        # --- 2. authors_ru_block
        authors_ru = []
        abstract_ru = []

        while i < n:
            p = clean_paragraphs[i]

            if self.is_keywords_ru_line(p) or self.is_keywords_en_line(p):
                break
            if self.is_reference_block(p):
                break
            if self.is_abstract_en_line(p):
                break
            if self.looks_like_english_title(p):
                break

            # явная аннотация
            if self.is_abstract_ru_line(p):
                line = p
                norm = self.normalize_text(line)

                if (
                    norm.startswith("аннотация.")
                    or norm.startswith("аннотация:")
                    or norm.startswith("аннотация ")
                ):
                    stripped = re.sub(r"^\s*аннотация(?:\s*[\.:]|\s+)\s*", "", line, flags=re.IGNORECASE).strip()
                    if stripped:
                        abstract_ru.append(stripped)
                i += 1
                break

            # сначала проверяем abstract-like
            if self.looks_like_abstract_paragraph(p):
                abstract_ru.append(p)
                i += 1
                break

            kind = self.classify_author_line(p)

            if kind in ("name", "spin_code", "org_text") or "@" in p:
                authors_ru.append(p)
                i += 1
                continue

            break

        if authors_ru:
            doc_struct["authors_ru_block"] = authors_ru

        # --- 3. abstract_ru continuation
        while i < n:
            p = clean_paragraphs[i]

            if self.is_keywords_ru_line(p):
                break
            if self.is_abstract_en_line(p) or self.is_keywords_en_line(p):
                break
            if self.is_reference_block(p):
                break
            if self.looks_like_funding_line(p):
                break
            if self.looks_like_english_title(p):
                break

            if self.looks_like_abstract_paragraph(p):
                abstract_ru.append(p)
                i += 1
                continue

            break

        if abstract_ru:
            doc_struct["abstract_ru"] = abstract_ru

            if abstract_ru:
                doc_struct["abstract_ru"] = abstract_ru

        # --- 4. keywords_ru
        if i < n and self.is_keywords_ru_line(clean_paragraphs[i]):
            doc_struct["keywords_ru"] = [clean_paragraphs[i]]
            i += 1

        # --- 5. english block
        title_en = None
        authors_en = []
        abstract_en = []
        keywords_en = None

        english_start = i

        # Если после RU keywords идет англ. title — не даем reference_block украсть его
        if english_start < n and self.looks_like_english_title(clean_paragraphs[english_start]):
            title_en = [clean_paragraphs[english_start]]
            i = english_start + 1

            # authors_en
            while i < n:
                p = clean_paragraphs[i]

                if self.is_abstract_en_line(p) or self.is_keywords_en_line(p):
                    break

                # если встретили явную литературу, но authors_en уже есть — заканчиваем блок
                if self.is_reference_block(p):
                    break

                if self.is_english_author_line(p):
                    authors_en.append(p)
                    i += 1
                    continue

                break

            # abstract_en
            if i < n and self.is_abstract_en_line(clean_paragraphs[i]):
                line = clean_paragraphs[i]
                norm = self.normalize_text(line)

                if norm.startswith("abstract.") or norm.startswith("abstract:"):
                    stripped = re.sub(r"^\s*abstract\s*[\.:]\s*", "", line, flags=re.IGNORECASE).strip()
                    if stripped:
                        abstract_en.append(stripped)
                    i += 1
                else:
                    i += 1
                    while i < n:
                        p = clean_paragraphs[i]
                        if self.is_keywords_en_line(p) or self.is_reference_block(p):
                            break
                        if self.looks_like_abstract_paragraph(p):
                            abstract_en.append(p)
                            i += 1
                            continue
                        break

            # keywords_en
            if i < n and self.is_keywords_en_line(clean_paragraphs[i]):
                keywords_en = [clean_paragraphs[i]]
                i += 1

        if title_en:
            doc_struct["title_en"] = title_en
        if authors_en:
            doc_struct["authors_en_block"] = authors_en
        if abstract_en:
            doc_struct["abstract_en"] = abstract_en
        if keywords_en:
            doc_struct["keywords_en"] = keywords_en

        # --- 6. body_block
        body = []
        while i < n:
            p = clean_paragraphs[i]
            if self.is_reference_block(p):
                break
            body.append(p)
            i += 1

        if body:
            doc_struct["body_block"] = body

        # --- 7. reference_block
            if i < n:
                ref_start = None
                for k in range(i, n):
                    p = clean_paragraphs[k]

                    if self.is_reference_block(p):
                        ref_start = k
                        break

                if ref_start is not None:
                    ref_end = self.find_reference_block_end(clean_paragraphs, ref_start)
                    doc_struct["reference_block"] = clean_paragraphs[ref_start:ref_end]
                    i = ref_end


            # --- 8. english block after russian references
                if i < n and doc_struct.get("title_en") in (None, []):
                    if self.looks_like_english_title(clean_paragraphs[i]):
                        title_en = [clean_paragraphs[i]]
                        i += 1

                        authors_en = []
                        while i < n:
                            p = clean_paragraphs[i]

                            if self.is_abstract_en_line(p) or self.is_keywords_en_line(p) or self.is_reference_block(p):
                                break

                            if self.is_english_author_line(p):
                                authors_en.append(p)
                                i += 1
                                continue

                            break

                        abstract_en = []
                        if i < n and self.is_abstract_en_line(clean_paragraphs[i]):
                            line = clean_paragraphs[i]
                            stripped = re.sub(r"^\s*(abstract|annotation)(?:\s*[\.:]|\s+)\s*", "", line, flags=re.IGNORECASE).strip()
                            if stripped:
                                abstract_en.append(stripped)
                            i += 1

                            while i < n:
                                p = clean_paragraphs[i]
                                if self.is_keywords_en_line(p) or self.is_reference_block(p):
                                    break
                                if self.looks_like_abstract_paragraph(p):
                                    abstract_en.append(p)
                                    i += 1
                                    continue
                                break

                        keywords_en = None
                        if i < n and self.is_keywords_en_line(clean_paragraphs[i]):
                            keywords_en = [clean_paragraphs[i]]
                            i += 1

                        doc_struct["title_en"] = title_en
                        if authors_en:
                            doc_struct["authors_en_block"] = authors_en
                        if abstract_en:
                            doc_struct["abstract_en"] = abstract_en
                        if keywords_en:
                            doc_struct["keywords_en"] = keywords_en            

        return doc_struct

    def looks_like_reference_item(self, text: str) -> bool:
        text = " ".join(text.replace("\xa0", " ").split()).strip()
        if not text:
            return False

        text_norm = text.lower()

        if re.match(r"^\d+\.\s+", text):
            return True
        if re.match(r"^\[\d+\]", text):
            return True
        if "//" in text:
            return True
        if re.search(r"\bdoi\b", text_norm):
            return True
        if re.search(r"\bvol\.?\b", text_norm):
            return True
        if re.search(r"\bno\.?\b", text_norm):
            return True
        if re.search(r"\bp\.?\s*\d", text_norm):
            return True
        if re.search(r"№\s*\d+", text):
            return True

        return False

    def find_reference_block_end(self, clean_paragraphs: list[str], start_idx: int) -> int:
        i = start_idx + 1
        n = len(clean_paragraphs)

        while i < n:
            p = clean_paragraphs[i]

            # если начался английский заголовок — стоп
            if self.looks_like_english_title(p):
                break

            # если пошла английская аннотация/keywords — тоже стоп
            if self.is_abstract_en_line(p) or self.is_keywords_en_line(p):
                break

            # если пошел английский author block без title
            if self.is_english_author_line(p):
                break

            i += 1

        return i

    def is_english_author_line(self, text: str) -> bool:
        text = " ".join(text.replace("\xa0", " ").split()).strip()
        if not text:
            return False

        text_norm = text.lower()

        # prose точно не авторы
        prose_starts = (
            "the article",
            "the paper",
            "this paper",
            "this article",
            "a generalization",
            "results of",
            "the study",
        )
        if any(text_norm.startswith(x) for x in prose_starts):
            return False

        if "@" in text:
            return True

        # инициалы + фамилия
        if re.search(r"\b[A-Z]\.[A-Z]?\.\s*[A-Z][a-z]+", text):
            return True

        # Имя Фамилия / First Middle Last
        if re.search(r"\b[A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,3}\b", text):
            return True

        # организация
        org_markers = (
            "university",
            "institute",
            "academy",
            "laboratory",
            "research center",
            "research institute",
            "moscow",
            "russia",
        )
        if any(marker in text_norm for marker in org_markers):
            return True

        if text_norm.startswith("spin-code"):
            return True

        return False

    def looks_like_english_title(self, text: str) -> bool:
        text = " ".join(text.replace("\xa0", " ").split()).strip()
        if not text:
            return False

        if self.is_reference_block(text):
            return False
        if self.is_keywords_en_line(text):
            return False
        if self.is_abstract_en_line(text):
            return False

        latin_words = re.findall(r"[A-Za-z]{4,}", text)
        cyr_words = re.findall(r"[А-Яа-яЁё]{3,}", text)

        if len(latin_words) >= 4 and len(latin_words) > len(cyr_words):
            return True

        return False

    
        


    def classify_author_line(self, text: str) -> str:
        text = " ".join(text.replace("\xa0", " ").split()).strip()
        if not text:
            return "unknown"

        if self.is_spin_line(text):
            return "spin_code"

        if self.is_author_ru_line(text) or self.is_author_en_line(text):
            return "name"

        if self.is_org_line_universal(text):
            return "org_text"

        if self.is_email_line(text):
            return "email"

        return "unknown"

    def load_cityies_set(self): # вызов в ините, если надо будет вжарить
        with open(CITYIES_PATH, "r", encoding="utf-8") as f:
            all_cityies = json.load(f)
        self.city_set_ru = {item["name"] for item in all_cityies}
        try:
            transltr = translator.YandexTranslator(source="ru", target="en", api_key=API_KEY)
            self.city_set_en = {transltr.translate(item) for item in self.city_set_ru}
        except Exception as e:
            self.city_set_en = {}

    def overload_attr(self): 
        return self.city_set_ru, self.city_set_en        

   
    def is_org_line_universal(self, text: str, from_set=False) -> bool:
        text = " ".join(text.replace("\xa0", " ").split()).strip()
        if not text:
            return False

        text_up = text.upper()

        org_markers = {
            "МГУ", "МГТУ", "МФТИ", "МИФИ", "МАИ", "МЭИ", "РАН",
            "УНИВЕРСИТЕТ", "УНИВЕРСИТЕТА", "ИНСТИТУТ", "ИНСТИТУТА",
            "АКАДЕМИЯ", "ЦЕНТР", "ЛАБОРАТОРИЯ",
            "UNIVERSITY", "INSTITUTE", "ACADEMY", "CENTER", "CENTRE", "LABORATORY",
            "FSBEI", "HE", "NRU", "RAS", "MSU", "BMSTU", "MPEI", "MAI", "MEPHI",
            "ГИДРОМЕТЦЕНТР", "РОСБИОТЕХ", "РУДН", "РУТ", "ИПМЕХ"
        }

        geo_markers = {
            "МОСКВА", "РОССИЯ", "САНКТ-ПЕТЕРБУРГ", "ЯРОСЛАВЛЬ",
            "MOSCOW", "RUSSIA", "ST. PETERSBURG", "ST PETERSBURG"
        }

        # строка начинается с индекса аффилиации и содержит маркер организации
        if re.match(r"^\d+\s*", text):
            if any(marker in text_up for marker in org_markers | geo_markers):
                return True

        if "," in text and any(marker in text_up for marker in geo_markers):
            return True

        if any(marker in text_up for marker in org_markers):
            return True

        return False
        

    def extract_phone(self, text: str) -> str | None:
        pattern = r"(?:\+?\d[\d\s\-\(\)]{9,}\d)"
        match = re.search(pattern, text)
        if match:
            return match.group(0).strip()
        return None

    def looks_like_funding_line(self, text: str) -> bool:
        text_norm = self.normalize_text(text)
        starts = (
            "исследование выполнено",
            "работа выполнена",
            "финансирование",
            "благодарности",
            "исследование выполнено в рамках",
            "the study was conducted",
            "funding",
            "acknowledgements",
        )
        return any(text_norm.startswith(x) for x in starts)

    def looks_like_abstract_paragraph(self, text: str) -> bool:
        text = " ".join(text.replace("\xa0", " ").split()).strip()
        if not text:
            return False

        text_norm = text.lower()

        bad_starts = (
            "удк",
            "ключевые слова",
            "keywords",
            "abstract",
            "аннотация",
            "введение",
            "методы",
            "методы и результаты",
            "результаты",
            "заключение",
            "финансирование",
            "благодарности",
            "список источников",
            "список литературы",
            "литература",
            "references",
        )

        if any(text_norm.startswith(x) for x in bad_starts):
            return False

        good_starts = (
            "представлены",
            "рассмотрены",
            "рассмотрено",
            "исследованы",
            "исследовано",
            "получены",
            "получено",
            "установлено",
            "выяснено",
            "предложен",
            "предложена",
            "предложены",
            "показано",
            "показаны",
            "the article presents",
            "the paper presents",
            "the paper considers",
            "results of",
        )

        if any(text_norm.startswith(x) for x in good_starts):
            return True

        if len(text) < 120 or text.count(" ") < 12:
            return False

        return True


    def struct_authors_block(self, authors_block: list[str]) -> dict:
        authors_info = []
        author = None

        for auth_item in authors_block:
            struct_item = self.classify_author_line(auth_item)

            if struct_item == "name":
                author = {
                    "name": auth_item,
                    "email": self.extract_email(auth_item),
                    "phone": self.extract_phone(auth_item),
                    "spin_code": None,
                    "org_ref": None,
                    "org_text": None,
                }
                authors_info.append(author)
                continue

            if author is None:
                continue

            if struct_item == "spin_code":
                author["spin_code"] = auth_item
            elif struct_item == "org_text":
                author["org_text"] = auth_item

        return {"authors": authors_info}
    
    # def recover_abstract_ru_without_marker(self, doc_struct: dict) -> dict:
    #     if doc_struct.get("abstract_ru") not in (None, []):
    #         return doc_struct

    #     authors_block = doc_struct.get("authors_ru_block") or []
    #     body_block = doc_struct.get("body_block") or []
    #     keywords_ru = doc_struct.get("keywords_ru") or []

    #     # кейс 1: после authors ничего не выделено, но body уже есть
    #     # и первый body-абзац выглядит как аннотация
    #     if body_block:
    #         first = body_block[0]

    #         # короткий/средний связный абзац - больше похож на аннотацию,
    #         # чем на основной текст статьи
    #         if len(first) <= 1200 and first.count(" ") >= 8:
    #             text_norm = self.normalize_text(first)

    #             bad_starts = (
    #                 "введение",
    #                 "методы",
    #                 "методы и результаты",
    #                 "заключение",
    #                 "финансирование",
    #                 "список источников",
    #                 "литература",
    #             )

    #             if not any(text_norm.startswith(x) for x in bad_starts):
    #                 doc_struct["abstract_ru"] = [first]
    #                 doc_struct["body_block"] = body_block[1:] if len(body_block) > 1 else None

    #     return doc_struct

    def postprocess_annot_doc(self, doc_struct: dict) -> dict:
        abs_block = doc_struct.get("abstract_ru") or []
        body_block = doc_struct.get("body_block") or []

        # спецкейс:
        # abstract_ru = ["АННОТАЦИЯ"]
        # body_block = [title, authors..., abstract...]
        if len(abs_block) == 1 and self.is_abstract_header_ru(abs_block[0]) and body_block:
            title_ru = None
            authors_ru = []
            abstract_ru = []

            # 1. первая строка body_block = title
            title_ru = [body_block[0]]

            # 2. дальше подряд идут author-строки
            idx = 1
            while idx < len(body_block):
                line = body_block[idx]
                kind = self.classify_author_line(line)

                if kind in ("name", "spin_code", "org_text"):
                    authors_ru.append(line)
                    idx += 1
                    continue

                break

            # 3. остаток = abstract
            abstract_ru = body_block[idx:]

            doc_struct["title_ru"] = title_ru if title_ru else None
            doc_struct["authors_ru_block"] = authors_ru if authors_ru else None
            doc_struct["abstract_ru"] = abstract_ru if abstract_ru else None

            # это не body статьи
            doc_struct["body_block"] = None

        return doc_struct
    
    def has_application_markers(self, clean_paragraphs: list[str]) -> bool:
        if not clean_paragraphs:
            return True

        joined = "\n".join(clean_paragraphs).lower()

        ignore_markers = (
            "заявка на участие",
            "заявка участника",
            "форма заявки",
            "регистрационная форма",
            "сведения об авторе",
            "сведения об участнике",
            "анкета участника",
            "экспертное заключение",
            "«ФУНДАМЕНТАЛЬНЫЕ И ПРИКЛАДНЫЕ ЗАДАЧИ МЕХАНИКИ»".lower(),
            "Fundamental and applied problems of mechanics (FAPM-2025)",
            "2-5 декабря 2025"
        )

        return any(marker in joined for marker in ignore_markers)

    def detect_document_kind(self, clean_paragraphs: list[str], doc_struct: dict) -> str:
        has_app = self.has_application_markers(clean_paragraphs)

        thesis_score = 0

        if doc_struct.get("title_ru") not in (None, "", []):
            thesis_score += 1
        if doc_struct.get("authors_ru_block") not in (None, "", []):
            thesis_score += 1
        if doc_struct.get("abstract_ru") not in (None, "", []):
            thesis_score += 1
        if doc_struct.get("keywords_ru") not in (None, "", []):
            thesis_score += 1
        if doc_struct.get("body_block") not in (None, "", []):
            thesis_score += 1

        # нормальные тезисы
        if thesis_score >= 3:
            if has_app:
                return "mixed"
            return "thesis"

        # просто заявка
        if has_app:
            return "application"

        return "unknown"
    
    def find_en_start_in_reference_block(self, ref_block: list[str]) -> int | None:
        for i, line in enumerate(ref_block):
            if self.looks_like_english_title(line):
                return i
        return None

    def split_polluted_reference_block(self, doc_struct: dict) -> dict:
        ref_block = doc_struct.get("reference_block") or []
        if not ref_block:
            return doc_struct

        en_start = self.find_en_start_in_reference_block(ref_block)
        if en_start is None:
            return doc_struct

        ru_ref = ref_block[:en_start]
        en_tail = ref_block[en_start:]

        doc_struct["reference_block"] = ru_ref if ru_ref else None

        i = 0
        n = len(en_tail)

        # --- title_en
        if i < n and self.looks_like_english_title(en_tail[i]):
            doc_struct["title_en"] = [en_tail[i]]
            i += 1

        # --- authors_en_block
        authors_en = []
        while i < n:
            p = en_tail[i]

            if self.is_abstract_en_line(p):
                break
            if self.is_keywords_en_line(p):
                break
            if self.normalize_text(p) == "references":
                break

            # ВАЖНО: сначала проверяем, не началась ли уже аннотация
            if self.looks_like_abstract_paragraph(p):
                break

            if self.looks_like_funding_line(p):
                break

            if self.is_english_author_line(p):
                authors_en.append(p)
                i += 1
                continue

            break

        if authors_en:
            doc_struct["authors_en_block"] = authors_en

        # --- abstract_en
        abstract_en = []
        if i < n:
            p = en_tail[i]
            norm = self.normalize_text(p)

            if self.is_abstract_en_line(p):
                stripped = re.sub(
                    r"^\s*(abstract|annotation)(?:\s*[\.:]|\s+)\s*",
                    "",
                    p,
                    flags=re.IGNORECASE
                ).strip()
                if stripped:
                    abstract_en.append(stripped)
                i += 1

            elif self.looks_like_abstract_paragraph(p):
                abstract_en.append(p)
                i += 1

            while i < n:
                p = en_tail[i]

                if self.is_keywords_en_line(p):
                    break
                if self.normalize_text(p) == "references":
                    break
                if self.looks_like_funding_line(p):
                    break

                if self.looks_like_abstract_paragraph(p):
                    abstract_en.append(p)
                    i += 1
                    continue

                break

        if abstract_en:
            doc_struct["abstract_en"] = abstract_en

        # --- keywords_en
        if i < n and self.is_keywords_en_line(en_tail[i]):
            doc_struct["keywords_en"] = [en_tail[i]]
            i += 1

        return doc_struct

    def debug_log(self, *parts, path=r"C:\Users\DanichMA.VASTA\Desktop\Python_training\ArticleNormalizer\app\workspace\debug_parser.txt"):
        with open(path, "a", encoding="utf-8") as f:
            f.write(" | ".join(map(str, parts)) + "\n")

    def get_parse_data(self, path):
        raw_data = self.read(path)
        if raw_data is None:
            return None

        clean_data = self.cleanText(raw_data["paragraphs"])
        clean_paragraphs = clean_data["clean_paragraphs"]

        doc_map_raw = self.find_all_marker_indexes(clean_paragraphs)
        doc_map_full = self.build_full_docmap(doc_map_raw)
        doc_struct = self.split_into_blocks(clean_paragraphs, doc_map_full, doc_map_raw)

        # нормализация inline abstract
        doc_struct = self.postprocess_inline_abstracts(doc_struct) if hasattr(self, "postprocess_inline_abstracts") else doc_struct
        doc_struct = self.cleanup_abstract_markers(doc_struct)

        # rescue EN from polluted references
        doc_struct = self.recover_en_from_reference_block(doc_struct)

        # abstract from body
        doc_struct = self.recover_abstract_from_body(doc_struct)

        # fallback parser
        fallback_struct = None
        if (
            doc_struct["title_ru"] is None
            or doc_struct["authors_ru_block"] is None
            or doc_struct["abstract_ru"] is None
        ):
            fallback_struct = self.build_fallback_doc_struct(clean_paragraphs)

            for key in ("title_ru", "authors_ru_block", "abstract_ru", "keywords_ru"):
                if doc_struct.get(key) in (None, []) and fallback_struct.get(key) not in (None, []):
                    doc_struct[key] = fallback_struct[key]

            for key in ("title_en", "authors_en_block", "abstract_en", "keywords_en", "body_block", "reference_block"):
                if doc_struct.get(key) in (None, []) and fallback_struct.get(key) not in (None, []):
                    doc_struct[key] = fallback_struct[key]

        # recover header from pre-abstract lines if still empty
        doc_struct = self.recover_header_from_pre_abstract_block(clean_paragraphs, doc_map_raw, doc_struct)

        # one more pass for abstract from fallback-body
        if fallback_struct is not None:
            doc_struct = self.merge_fallback_abstract_from_body(doc_struct, fallback_struct)

        doc_struct = self.cleanup_abstract_markers(doc_struct)          

        doc_kind = self.detect_document_kind(clean_paragraphs, doc_struct)
        if doc_kind == "application":
            return None

        doc_struct["_doc_kind"] = doc_kind
        return doc_struct

class DocxValidator:

    def __init__(self):
        pass

    def validate_doc_struct(self, doc_struct: dict) -> dict:
        report = {
            "status": None,
            "issues": [],
            "optional_miss": []
        }

        if doc_struct is None:
            report["status"] = "ignored"
            report["_doc_kind"] =  "application"
            return report

        title_ok = doc_struct.get("title_ru") not in (None, "", [])
        authors_ok = doc_struct.get("authors_ru_block") not in (None, "", [])
        abstract_ok = doc_struct.get("abstract_ru") not in (None, "", [])
        body_ok = doc_struct.get("body_block") not in (None, "", [])

        optional = (
            "udk",
            "keywords_ru",
            "title_en",
            "authors_en_block",
            "abstract_en",
            "keywords_en",
            "reference_block",
        )

        if not title_ok:
            report["issues"].append("missing_title_ru")
        if not authors_ok:
            report["issues"].append("missing_authors_ru_block")

        if title_ok and authors_ok and abstract_ok:
            report["status"] = "ok"
        elif title_ok and abstract_ok and not authors_ok:
            report["status"] = "partial"
            report["issues"].append("missing_authors_ru_block")    
        elif title_ok and authors_ok and body_ok:
            report["status"] = "partial"
            report["issues"].append("missing_abstract_ru")
        else:
            if not abstract_ok:
                report["issues"].append("missing_abstract_ru")
            report["status"] = "invalid"

        for opt in optional:
            if doc_struct.get(opt) in (None, "", []):
                report["optional_miss"].append(f"missing_{opt}")

        report["issues"] = list(dict.fromkeys(report["issues"]))
        report["optional_miss"] = list(dict.fromkeys(report["optional_miss"]))        

        return report          
                
                        


                





            








        
    



































































        



