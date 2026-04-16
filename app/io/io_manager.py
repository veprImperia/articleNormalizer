# io_manager.py
import json
import shutil
from pathlib import Path


class InputCollector:
    """
    Сборщик входных документов с предварительной фильтрацией по пакетам подачи.

    Логика:
    - обходит все подпапки input_root
    - рассматривает каждую папку как потенциальный "пакет подачи"
    - если в папке есть признаки экспертного пакета, docx копируются в collected_dir
    - если признаков нет, docx из папки не копируются, папка уходит в quarantine
    - формируется протокол отбора
    """

    EXPERT_NAME_MARKERS = (
        "экс",
        "эксп",
        "эксперт",
        "экспертное",
        "закл",
        "заключ",
        "заявка",
    )

    EXPERT_EXTENSIONS = {
        ".pdf",
        ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".gif", ".webp",
    }

    def __init__(self, input_root: Path, collected_dir: Path, quarantine_dir: Path | None = None):
        self.root = input_root
        if not self.root.exists():
            raise FileNotFoundError("Указанный путь не существует")

        self.collected_dir = collected_dir
        self.collected_dir.mkdir(parents=True, exist_ok=True)

        self.quarantine_dir = quarantine_dir
        if self.quarantine_dir is not None:
            self.quarantine_dir.mkdir(parents=True, exist_ok=True)

        self.copiedFiles = []

        self.protocol = {
            "accepted_packages": [],
            "quarantined_packages": [],
            "copied_docx": [],
            "skipped_duplicates": [],
        }

    # ---------------------------------------------------------
    # PUBLIC API
    # ---------------------------------------------------------

    def collect_files(self):
        package_dirs = self._find_package_dirs()

        if not package_dirs:
            return

        for folder in package_dirs:
            self._process_package(folder)

    def get_collected_files(self):
        self.copiedFiles = list(self.collected_dir.rglob("*.docx"))
        return self.copiedFiles

    def get_protocol(self) -> dict:
        return self.protocol

    def save_protocol(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.protocol, f, indent=2, ensure_ascii=False)

    def save_human_report(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)

        lines = []
        lines.append("INPUT COLLECTOR REPORT")
        lines.append("=" * 80)
        lines.append(f"ROOT: {self.root}")
        lines.append(f"COLLECTED DIR: {self.collected_dir}")
        if self.quarantine_dir is not None:
            lines.append(f"QUARANTINE DIR: {self.quarantine_dir}")
        lines.append("")

        accepted = self.protocol.get("accepted_packages", [])
        quarantined = self.protocol.get("quarantined_packages", [])
        copied = self.protocol.get("copied_docx", [])
        skipped = self.protocol.get("skipped_duplicates", [])

        lines.append("SUMMARY")
        lines.append("-" * 80)
        lines.append(f"Accepted packages:   {len(accepted)}")
        lines.append(f"Quarantined packages:{len(quarantined)}")
        lines.append(f"Copied docx:         {len(copied)}")
        lines.append(f"Skipped duplicates:  {len(skipped)}")
        lines.append("")

        lines.append("ACCEPTED PACKAGES")
        lines.append("-" * 80)
        if accepted:
            for idx, item in enumerate(accepted, start=1):
                lines.append(f"{idx}. Folder: {item['folder']}")
                lines.append("   DOCX:")
                for name in item.get("docx_files", []):
                    lines.append(f"     - {name}")
                lines.append("   Expert signs:")
                for sign in item.get("expert_signs", []):
                    lines.append(f"     - {sign}")
                lines.append("")
        else:
            lines.append("No accepted packages.")
            lines.append("")

        lines.append("QUARANTINED PACKAGES")
        lines.append("-" * 80)
        if quarantined:
            for idx, item in enumerate(quarantined, start=1):
                lines.append(f"{idx}. Folder: {item['folder']}")
                lines.append(f"   Reason: {item.get('reason', 'unknown')}")
                lines.append("   DOCX:")
                for name in item.get("docx_files", []):
                    lines.append(f"     - {name}")
                lines.append("")
        else:
            lines.append("No quarantined packages.")
            lines.append("")

        lines.append("SKIPPED DUPLICATES / COPY ERRORS")
        lines.append("-" * 80)
        if skipped:
            for idx, item in enumerate(skipped, start=1):
                lines.append(f"{idx}. Source: {item.get('source')}")
                lines.append(f"   Reason: {item.get('reason')}")
                lines.append("")
        else:
            lines.append("No skipped duplicates or copy errors.")
            lines.append("")

        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))        

    # ---------------------------------------------------------
    # INTERNAL
    # ---------------------------------------------------------

    def _find_package_dirs(self) -> list[Path]:
        """
        Пакет = папка, в которой есть хотя бы один docx.
        """
        result = []

        for folder in self.root.rglob("*"):
            if not folder.is_dir():
                continue

            if self.collected_dir == folder or self.collected_dir in folder.parents:
                continue

            if self.quarantine_dir is not None and (self.quarantine_dir == folder or self.quarantine_dir in folder.parents):
                continue

            docx_files = [
                f for f in folder.iterdir()
                if f.is_file() and f.suffix.lower() == ".docx" and not f.name.startswith("~$")
            ]

            if docx_files:
                result.append(folder)

        return result

    def _process_package(self, folder: Path):
        files = [f for f in folder.iterdir() if f.is_file()]
        docx_files = [
            f for f in files
            if f.suffix.lower() == ".docx" and not f.name.startswith("~$")
        ]

        if not docx_files:
            return

        has_expert_sign = self._has_expert_sign(files)

        if has_expert_sign:
            accepted_entry = {
                "folder": str(folder),
                "docx_files": [f.name for f in docx_files],
                "expert_signs": self._describe_expert_signs(files),
            }
            self.protocol["accepted_packages"].append(accepted_entry)

            for file in docx_files:
                self._copy_docx(file)

        else:
            quarantine_entry = {
                "folder": str(folder),
                "docx_files": [f.name for f in docx_files],
                "reason": "missing expert signs: no pdf, no image files, no filename markers (эксп/закл/заявка)",
            }
            self.protocol["quarantined_packages"].append(quarantine_entry)

            if self.quarantine_dir is not None:
                self._copy_to_quarantine(folder, docx_files)

    def _has_expert_sign(self, files: list[Path]) -> bool:
        for f in files:
            name_low = f.name.lower()
            suffix_low = f.suffix.lower()

            if any(marker in name_low for marker in self.EXPERT_NAME_MARKERS):
                return True

            if suffix_low in self.EXPERT_EXTENSIONS:
                return True

        return False

    def _describe_expert_signs(self, files: list[Path]) -> list[str]:
        found = []

        for f in files:
            name_low = f.name.lower()
            suffix_low = f.suffix.lower()

            matched_markers = [m for m in self.EXPERT_NAME_MARKERS if m in name_low]
            if matched_markers:
                found.append(f"name:{f.name} -> {matched_markers}")

            if suffix_low in self.EXPERT_EXTENSIONS:
                found.append(f"ext:{f.name}")

        return found

    def _copy_docx(self, file: Path):
        base_target = self.collected_dir / file.name
        if base_target.exists():
            self.protocol["skipped_duplicates"].append({
                "source": str(file),
                "reason": "target_exists",
            })
            return

        target = self._make_unique_target(file)

        try:
            shutil.copy(file, target)
            self.protocol["copied_docx"].append({
                "source": str(file),
                "target": str(target),
            })
        except Exception as e:
            self.protocol["skipped_duplicates"].append({
                "source": str(file),
                "reason": f"copy_error: {e}",
            })

    def _copy_to_quarantine(self, folder: Path, docx_files: list[Path]):
        folder_target = self.quarantine_dir / folder.name
        folder_target.mkdir(parents=True, exist_ok=True)

        for file in docx_files:
            target = folder_target / file.name
            try:
                if not target.exists():
                    shutil.copy(file, target)
            except Exception:
                pass

    def _make_unique_target(self, file: Path):
        target = self.collected_dir / file.name
        if not target.exists():
            return target

        counter = 1
        while True:
            new_name = Path(f"{file.stem}_{counter}{file.suffix}")
            target = self.collected_dir / new_name
            if not target.exists():
                return target
            counter += 1