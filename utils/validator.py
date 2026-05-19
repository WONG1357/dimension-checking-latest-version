from __future__ import annotations

from collections import defaultdict
from typing import Any

import pandas as pd


def build_summary(detail_df: pd.DataFrame, sheet_summary_df: pd.DataFrame, dimension_summary_df: pd.DataFrame) -> dict[str, Any]:
    """Build dashboard summaries from the extracted validation table."""
    total_records = int(sheet_summary_df["inspection_record_count"].sum()) if not sheet_summary_df.empty else 0
    total_measurements = len(detail_df)
    pass_count = int((detail_df["status"] == "PASS").sum()) if not detail_df.empty else 0
    fail_count = int((detail_df["status"] == "FAIL").sum()) if not detail_df.empty else 0
    spec_missing_count = int((detail_df["status"] == "SPEC_MISSING").sum()) if not detail_df.empty else 0

    overall_result = "FAIL" if (fail_count or spec_missing_count) else "PASS" if total_measurements else "NO DATA"

    fail_by_dimension = (
        detail_df.loc[detail_df["status"] == "FAIL", "dimension_unique_name"]
        .value_counts()
        .rename_axis("dimension_unique_name")
        .reset_index(name="fail_count")
        if not detail_df.empty
        else pd.DataFrame(columns=["dimension_unique_name", "fail_count"])
    )
    pass_fail_counts = pd.DataFrame(
        {
            "status": ["PASS", "FAIL", "SPEC_MISSING"],
            "count": [pass_count, fail_count, spec_missing_count],
        }
    )

    return {
        "overall": {
            "total_sheets": int(sheet_summary_df["sheet_name"].nunique()) if not sheet_summary_df.empty else 0,
            "total_records": total_records,
            "total_measurements": total_measurements,
            "pass_count": pass_count,
            "fail_count": fail_count,
            "spec_missing_count": spec_missing_count,
            "overall_result": overall_result,
        },
        "pass_fail_counts": pass_fail_counts,
        "fail_by_dimension": fail_by_dimension,
        "sheet_summary": sheet_summary_df,
        "dimension_summary": dimension_summary_df,
    }
