"""Re-export правил количеств из compare (без дублирования логики)."""

from .compare import (
    dedupe_quantity_rows,
    has_non_numeric_issue,
    infer_operator,
    is_false_numeric_mismatch,
    normalize_photo_row,
    refine_comparison_rows,
    normalize_quantity_row,
    quantity_requirement_met,
    should_include_comment_row,
)

__all__ = [
    "infer_operator",
    "quantity_requirement_met",
    "has_non_numeric_issue",
    "is_false_numeric_mismatch",
    "normalize_quantity_row",
    "normalize_photo_row",
    "dedupe_quantity_rows",
    "refine_comparison_rows",
    "should_include_comment_row",
]
