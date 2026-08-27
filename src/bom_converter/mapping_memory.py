from __future__ import annotations

import json
import os
import tempfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .models import ColumnMappingConfig, NameGenerationRule
from .text_utils import normalize_header


SCHEMA_VERSION = 2


@dataclass(frozen=True)
class SimilarRule:
    fingerprint: str
    similarity: float
    added_headers: list[str]
    missing_headers: list[str]


class MappingRuleStore:
    """Local-only rules. It never stores sample values, rows, IDs or images."""

    def __init__(self, path: str | Path):
        self.path = Path(path).resolve()

    @staticmethod
    def _empty() -> dict[str, Any]:
        return {
            "schema_version": SCHEMA_VERSION,
            "privacy": {
                "stores_headers_and_rules_only": True,
                "stores_sample_values": False,
                "stores_bom_rows": False,
                "stores_images": False,
                "stores_part_names_or_numbers": False,
            },
            "rules": {},
            "translation_terms": {},
        }

    def _load(self) -> dict[str, Any]:
        if not self.path.exists():
            return self._empty()
        data = json.loads(self.path.read_text(encoding="utf-8"))
        if not isinstance(data, dict) or data.get("schema_version") not in {1, SCHEMA_VERSION}:
            raise ValueError("Unsupported mapping rule database format")
        if not isinstance(data.get("rules"), dict):
            raise ValueError("Mapping rule database does not contain a rules object")
        if data.get("schema_version") == 1:
            data["schema_version"] = SCHEMA_VERSION
            privacy = data.setdefault("privacy", {})
            privacy.update({
                "stores_headers_and_rules_only": True,
                "stores_sample_values": False,
                "stores_bom_rows": False,
                "stores_images": False,
                "stores_part_names_or_numbers": False,
            })
        return data

    def _save(self, data: dict[str, Any]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        handle = tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            prefix="mapping_rules_",
            suffix=".tmp",
            dir=self.path.parent,
            delete=False,
        )
        temp_path = Path(handle.name)
        try:
            with handle:
                json.dump(data, handle, ensure_ascii=False, indent=2, sort_keys=True)
            os.replace(temp_path, self.path)
        except Exception:
            temp_path.unlink(missing_ok=True)
            raise

    def get_rule(self, workbook_fingerprint: str) -> dict[str, Any] | None:
        rule = self._load()["rules"].get(workbook_fingerprint)
        return dict(rule) if rule else None

    def recall(self, workbook_fingerprint: str) -> dict[int, str]:
        rule = self.get_rule(workbook_fingerprint)
        if not rule:
            return {}
        return {
            int(column): field
            for column, field in rule.get("mappings", {}).items()
            if field
        }

    def recall_configs(self, workbook_fingerprint: str) -> dict[int, ColumnMappingConfig]:
        rule = self.get_rule(workbook_fingerprint)
        if not rule:
            return {}
        result: dict[int, ColumnMappingConfig] = {}
        for column, payload in rule.get("column_configs", {}).items():
            if not isinstance(payload, dict):
                continue
            source_col = int(column)
            result[source_col] = ColumnMappingConfig(
                source_col=source_col,
                process_type=str(payload.get("process_type") or "direct"),
                universal_field=payload.get("universal_field"),
                unit=payload.get("unit"),
                image_slot=payload.get("image_slot"),
                level_group=payload.get("level_group"),
                level_value=payload.get("level_value"),
                default_value=payload.get("default_value"),
            )
        return result

    def recall_name_rule(self, workbook_fingerprint: str) -> NameGenerationRule | None:
        rule = self.get_rule(workbook_fingerprint)
        if not rule:
            return None
        payload = rule.get("name_generation")
        if not isinstance(payload, dict):
            # Version 1/2 rules used this exact implicit behavior.
            return NameGenerationRule()

        def column(name: str) -> int | None:
            value = payload.get(name)
            return int(value) if value not in {None, ""} else None

        strategy = str(payload.get("strategy") or "fallback")
        if strategy not in {"fallback", "replace", "append"}:
            strategy = "fallback"
        return NameGenerationRule(
            strategy=strategy,
            original_name_col=column("original_name_col"),
            standard_name_col=column("standard_name_col"),
            gb_name_col=column("gb_name_col"),
            spec_col=column("spec_col"),
            template=str(payload.get("template") or "{名称} {GB} {规格}"),
            deduplicate=bool(payload.get("deduplicate", True)),
        )

    def save_rule(
        self,
        workbook_fingerprint: str,
        *,
        header_rows: list[int],
        headers: list[str],
        mappings: dict[int, str | None],
        units: dict[str, str] | None = None,
        ignored_columns: list[int] | None = None,
        sheet_name: str | None = None,
        header_start_row: int | None = None,
        header_end_row: int | None = None,
        data_start_row: int | None = None,
        header_paths: list[list[str]] | None = None,
        merged_structure_summary: list[str] | None = None,
        column_configs: dict[int, ColumnMappingConfig] | None = None,
        name_rule: NameGenerationRule | None = None,
    ) -> None:
        data = self._load()
        normalized_headers = [normalize_header(header) for header in headers]
        ignored = sorted(set(ignored_columns or [column for column, field in mappings.items() if not field]))
        serialized_configs: dict[str, dict[str, Any]] = {}
        for column, config in sorted((column_configs or {}).items()):
            serialized_configs[str(column)] = {
                "source_col": config.source_col,
                "process_type": config.process_type,
                "universal_field": config.universal_field,
                "unit": config.unit,
                "image_slot": config.image_slot,
                "level_group": config.level_group,
                "level_value": config.level_value,
                "default_value": config.default_value,
            }
        normalized_paths = [
            [normalize_header(piece) for piece in path if normalize_header(piece)]
            for path in (header_paths or [[header] for header in headers])
        ]
        effective_name_rule = name_rule or NameGenerationRule()
        data["rules"][workbook_fingerprint] = {
            "sheet_name": sheet_name,
            "header_rows": list(header_rows),
            "header_start_row": header_start_row if header_start_row is not None else (min(header_rows) if header_rows else None),
            "header_end_row": header_end_row if header_end_row is not None else (max(header_rows) if header_rows else None),
            "data_start_row": data_start_row,
            "normalized_headers": normalized_headers,
            "normalized_header_paths": normalized_paths,
            "merged_structure_summary": sorted(merged_structure_summary or []),
            "mappings": {
                str(column): field
                for column, field in sorted(mappings.items())
                if field
            },
            "column_configs": serialized_configs,
            "units": dict(sorted((units or {}).items())),
            "ignored_columns": ignored,
            "name_generation": {
                "strategy": effective_name_rule.strategy,
                "original_name_col": effective_name_rule.original_name_col,
                "standard_name_col": effective_name_rule.standard_name_col,
                "gb_name_col": effective_name_rule.gb_name_col,
                "spec_col": effective_name_rule.spec_col,
                "template": effective_name_rule.template,
                "deduplicate": effective_name_rule.deduplicate,
            },
            "updated_at_utc": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
        }
        self._save(data)

    def remember(self, workbook_fingerprint: str, mappings: dict[int, str]) -> None:
        """Compatibility helper for tests and callers that only have a fingerprint."""
        self.save_rule(
            workbook_fingerprint,
            header_rows=[],
            headers=[],
            mappings=mappings,
        )

    def find_similar(self, headers: list[str], minimum: float = 0.60) -> SimilarRule | None:
        current = {normalize_header(header) for header in headers if normalize_header(header)}
        best: SimilarRule | None = None
        for fingerprint, rule in self._load()["rules"].items():
            saved = {header for header in rule.get("normalized_headers", []) if header}
            if not saved:
                saved = {
                    "/".join(path)
                    for path in rule.get("normalized_header_paths", [])
                    if path
                }
            if not current or not saved:
                continue
            similarity = len(current & saved) / len(current | saved)
            candidate = SimilarRule(
                fingerprint,
                similarity,
                sorted(current - saved),
                sorted(saved - current),
            )
            if similarity >= minimum and (best is None or similarity > best.similarity):
                best = candidate
        return best


# Backward-compatible name used by the first memory tests.
MappingMemory = MappingRuleStore
