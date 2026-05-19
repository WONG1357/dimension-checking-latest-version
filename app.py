from __future__ import annotations

from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st

from utils.parser import parse_workbook
from utils.report_writer import export_validation_report
from utils.validator import build_summary


st.set_page_config(page_title="Injection In-Process Inspection Checker", layout="wide")
st.title("Injection In-Process Inspection Checker")
st.caption(
    "Upload an Excel inspection report to automatically validate all sample measurements against specification limits."
)


@st.cache_data(show_spinner=False)
def _parse_uploaded_file(file_bytes: bytes, file_name: str, parser_version: str):
    buffer = BytesIO(file_bytes)
    buffer.name = file_name
    return parse_workbook(buffer)


PARSER_VERSION = "metadata-box-v2"


st.subheader("Step 1: Upload Excel File")
st.write("Please upload your inspection Excel file to start validation.")
uploaded_file = st.file_uploader(
    "Upload inspection Excel file",
    type=["xlsx", "xlsm", "xls"],
)

if uploaded_file is None:
    st.info("No file uploaded yet. Please upload an Excel inspection workbook using the uploader above.")
    st.stop()

raw_bytes = uploaded_file.getvalue()

if uploaded_file.name.lower().endswith(".xls"):
    st.error("This dashboard currently supports .xlsx and .xlsm files. Please convert the .xls file to .xlsx or .xlsm before uploading.")
    st.stop()

try:
    parsed = _parse_uploaded_file(raw_bytes, uploaded_file.name, PARSER_VERSION)
except Exception:
    st.error("The file could not be processed. Please check that it is a valid inspection workbook and try again.")
    st.stop()

detail_df = pd.DataFrame(parsed["detail_rows"], columns=parsed["detail_df_columns"])
sheet_summary_df = pd.DataFrame(parsed["sheet_summaries"])
dimension_summary_df = pd.DataFrame(parsed["dimension_summaries"])

summary = build_summary(detail_df, sheet_summary_df, dimension_summary_df)

with st.sidebar:
    st.header("Filters")
    sheet_options = ["All Sheets"]
    if not sheet_summary_df.empty:
        sheet_options.extend(sorted(sheet_summary_df["sheet_name"].dropna().unique().tolist()))
    selected_sheet = st.selectbox("Sheet", sheet_options, index=0)
    status_options = ["All", "PASS", "FAIL", "SPEC_MISSING"]
    selected_status = st.selectbox("Status", status_options, index=0)
    dimension_options = ["All Dimensions"]
    if not detail_df.empty:
        dimension_options.extend(sorted(detail_df["dimension_unique_name"].dropna().unique().tolist()))
    selected_dimension = st.selectbox("Dimension", dimension_options, index=0)

filtered_df = detail_df.copy()
if selected_sheet != "All Sheets":
    filtered_df = filtered_df[filtered_df["sheet_name"] == selected_sheet]
if selected_status != "All":
    filtered_df = filtered_df[filtered_df["status"] == selected_status]
if selected_dimension != "All Dimensions":
    filtered_df = filtered_df[filtered_df["dimension_unique_name"] == selected_dimension]

st.success("File processed successfully.")

st.subheader("Step 2: Review Validation Results")
st.write(
    f"Uploaded file: {uploaded_file.name} | File size: {uploaded_file.size:,} bytes"
)

summary_col1, summary_col2, summary_col3, summary_col4, summary_col5 = st.columns(5)
summary_col1.metric("Sheets", summary["overall"]["total_sheets"])
summary_col2.metric("Measurements", summary["overall"]["total_measurements"])
summary_col3.metric("PASS", summary["overall"]["pass_count"])
summary_col4.metric("FAIL", summary["overall"]["fail_count"])
summary_col5.metric("Overall Result", summary["overall"]["overall_result"])

st.write("Sheet names:", ", ".join(sheet_summary_df["sheet_name"].astype(str).tolist()) if not sheet_summary_df.empty else "None detected")

if parsed["warnings"]:
    with st.expander("Parsing Warnings", expanded=False):
        for warning in parsed["warnings"]:
            st.warning(warning)

results_tab, sheets_tab, charts_tab = st.tabs(["Validation Table", "Workbook Summary", "Charts"])

with results_tab:
    if filtered_df.empty:
        st.info("No rows match the current filters.")
    else:
        display_cols = [
            "sheet_name",
            "inspection_time",
            "dimension_display_name",
            "nominal_value",
            "calculated_lower_limit",
            "calculated_upper_limit",
            "measurement_value",
            "status",
            "source_cell",
        ]
        display_df = filtered_df[display_cols].copy()
        styled = display_df.style.apply(
            lambda row: [
                "background-color: #F4CCCC; color: #000000"
                if row["status"] == "FAIL"
                else "background-color: #E2F0D9; color: #000000"
                if row["status"] == "PASS"
                else "background-color: #FFF2CC; color: #000000"
                for _ in row
            ],
            axis=1,
        )
        st.dataframe(styled, use_container_width=True, hide_index=True)

    if not filtered_df.empty:
        csv_data = filtered_df.to_csv(index=False).encode("utf-8-sig")
        excel_data = export_validation_report(
            detail_df=detail_df,
            sheet_summary_df=sheet_summary_df,
            overall_summary=summary["overall"],
            dimension_summary_df=dimension_summary_df,
        )
        download_col1, download_col2 = st.columns(2)
        download_col1.download_button(
            "Download CSV Report",
            data=csv_data,
            file_name="inspection_validation_report.csv",
            mime="text/csv",
        )
        download_col2.download_button(
            "Download Excel Report",
            data=excel_data,
            file_name="inspection_validation_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with sheets_tab:
    if not sheet_summary_df.empty:
        st.dataframe(sheet_summary_df, use_container_width=True, hide_index=True)
    else:
        st.info("No sheet summary available.")

with charts_tab:
    chart_col1, chart_col2 = st.columns(2)
    with chart_col1:
        pass_fail_counts = summary["pass_fail_counts"]
        if not pass_fail_counts.empty:
            fig = px.bar(pass_fail_counts, x="status", y="count", color="status", title="PASS vs FAIL")
            st.plotly_chart(fig, use_container_width=True)
    with chart_col2:
        fail_by_dimension = summary["fail_by_dimension"]
        if not fail_by_dimension.empty:
            fig = px.bar(
                fail_by_dimension.head(15),
                x="dimension_unique_name",
                y="fail_count",
                title="FAIL Count by Dimension",
            )
            st.plotly_chart(fig, use_container_width=True)

    trend_source = filtered_df.copy()
    if selected_dimension != "All Dimensions" and not trend_source.empty:
        trend_source = trend_source[trend_source["dimension_unique_name"] == selected_dimension]
    elif not trend_source.empty and trend_source["dimension_unique_name"].nunique() > 0:
        trend_source = trend_source[trend_source["dimension_unique_name"] == trend_source["dimension_unique_name"].iloc[0]]

    if trend_source.empty:
        st.info("Select a dimension with data to view the trend chart.")
    else:
        def _time_sort_key(value: str) -> int:
            try:
                hour, minute = value.split(":")[:2]
                return int(hour) * 60 + int(minute)
            except Exception:
                return 0

        trend_source = trend_source.copy()
        trend_source["time_sort"] = trend_source["inspection_time"].map(_time_sort_key)
        trend_source = trend_source.sort_values("time_sort")
        fig = px.line(
            trend_source,
            x="inspection_time",
            y="measurement_value",
            color="sheet_name",
            markers=True,
            title="Measurement Trend by Inspection Time",
        )
        st.plotly_chart(fig, use_container_width=True)
