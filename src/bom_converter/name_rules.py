from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Any, Callable, Iterable

from .models import ColumnMappingConfig, MappingDecision, NameGenerationRule


STRATEGY_LABELS = {
    "fallback": "原名称为空时组合",
    "replace": "始终使用组合名称",
    "append": "在原名称后追加",
}

TEMPLATE_PRESETS = {
    "空格连接": "{名称} {GB} {规格}",
    "括号格式": "{名称}（{GB}）{规格}",
    "横线格式": "{名称}-{GB}-{规格}",
    "自定义": "",
}

PLACEHOLDER_ALIASES = {
    "原名称": "original",
    "名称": "standard",
    "标准件名称": "standard",
    "GB": "gb",
    "GB名称": "gb",
    "规格": "spec",
}

DEFAULT_NAME_RULE = NameGenerationRule()


@dataclass(frozen=True)
class NameGenerationResult:
    final_name: str | None
    original_name: str | None
    combined_name: str | None
    removed_duplicates: list[str]


def clean_name_text(value: Any) -> str | None:
    """Clean display names without changing ×, Chinese brackets or other meaning."""

    if value is None:
        return None
    text = re.sub(r"[ \t]+", " ", str(value).strip())
    if text.casefold() in {"", "-", "—", "–", "n/a", "na", "none", "null", "无"}:
        return None
    return text


def _dedup_key(value: str) -> str:
    return re.sub(r"\s+", "", value).casefold()


def _deduplicate_fields(
    fields: list[tuple[str, str | None]],
    enabled: bool,
) -> tuple[dict[str, str | None], list[str]]:
    result: dict[str, str | None] = {}
    seen: dict[str, str] = {}
    removed: list[str] = []
    for label, value in fields:
        if not value:
            result[label] = None
            continue
        key = _dedup_key(value)
        if enabled and key in seen:
            removed.append(f"{label}“{value}”与{seen[key]}重复")
            result[label] = None
            continue
        seen[key] = label
        result[label] = value
    return result, removed


def _clean_rendered_name(value: str) -> str | None:
    value = re.sub(r"（\s*）|\(\s*\)|【\s*】|\[\s*\]", "", value)
    value = re.sub(r"\s*[-—–_/]+\s*", "-", value)
    value = re.sub(r"-{2,}", "-", value).strip(" -—–_/")
    value = re.sub(r"\s+", " ", value).strip()
    return value or None


def render_name_template(
    template: str,
    *,
    original: str | None,
    standard: str | None,
    gb: str | None,
    spec: str | None,
) -> str | None:
    values = {"original": original, "standard": standard, "gb": gb, "spec": spec}

    def replace(match: re.Match[str]) -> str:
        alias = PLACEHOLDER_ALIASES.get(match.group(1))
        if alias is None:
            return match.group(0)
        return values.get(alias) or ""

    return _clean_rendered_name(re.sub(r"\{([^{}]+)\}", replace, template))


def generate_name(
    rule: NameGenerationRule | None,
    *,
    original: Any = None,
    standard: Any = None,
    gb: Any = None,
    spec: Any = None,
) -> NameGenerationResult:
    rule = rule or DEFAULT_NAME_RULE
    original_text = clean_name_text(original)
    fields, removed = _deduplicate_fields(
        [
            ("原名称", original_text),
            ("标准件名称", clean_name_text(standard)),
            ("GB名称", clean_name_text(gb)),
            ("规格", clean_name_text(spec)),
        ],
        rule.deduplicate,
    )
    original_for_output = fields["原名称"] or original_text
    combined = render_name_template(
        rule.template or DEFAULT_NAME_RULE.template,
        original=fields["原名称"],
        standard=fields["标准件名称"],
        gb=fields["GB名称"],
        spec=fields["规格"],
    )
    if rule.strategy == "replace":
        final = combined
    elif rule.strategy == "append":
        if original_for_output and combined and _dedup_key(original_for_output) != _dedup_key(combined):
            final = f"{original_for_output} {combined}"
        else:
            final = original_for_output or combined
    else:
        final = original_for_output or combined
    return NameGenerationResult(_clean_rendered_name(final or ""), original_text, combined, removed)


def infer_name_rule_columns(
    decisions: Iterable[MappingDecision],
    configs: dict[int, ColumnMappingConfig] | None = None,
    base: NameGenerationRule | None = None,
) -> NameGenerationRule:
    base = base or NameGenerationRule()
    by_field: dict[str, int] = {}
    for decision in decisions:
        config = configs.get(decision.source_col) if configs else None
        field = config.universal_field if config else decision.universal_field
        if field and field not in by_field:
            by_field[field] = decision.source_col
    return NameGenerationRule(
        strategy=base.strategy,
        original_name_col=base.original_name_col or by_field.get("part_name"),
        standard_name_col=base.standard_name_col or by_field.get("standard_name"),
        gb_name_col=base.gb_name_col or by_field.get("standard_name_gb"),
        spec_col=base.spec_col or by_field.get("spec"),
        template=base.template or DEFAULT_NAME_RULE.template,
        deduplicate=base.deduplicate,
    )


def generate_name_from_columns(
    rule: NameGenerationRule | None,
    value_for_column: Callable[[int], Any],
    *,
    fallback_original: Any = None,
    fallback_standard: Any = None,
    fallback_gb: Any = None,
    fallback_spec: Any = None,
) -> NameGenerationResult:
    rule = rule or DEFAULT_NAME_RULE

    def value(column: int | None, fallback: Any) -> Any:
        return value_for_column(column) if column else fallback

    return generate_name(
        rule,
        original=value(rule.original_name_col, fallback_original),
        standard=value(rule.standard_name_col, fallback_standard),
        gb=value(rule.gb_name_col, fallback_gb),
        spec=value(rule.spec_col, fallback_spec),
    )
