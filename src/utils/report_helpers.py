"""
Report Helpers

Helper classes and functions for building Excel reports.
"""

import os
import glob
from typing import List, Dict, Any, Callable, Optional

from core.logging_config import get_logger

logger = get_logger(__name__)


class ReportBuilder:
    """Helper for building Excel reports."""

    def __init__(self) -> None:
        from modules.excel_automation import WorkbookHandler, ChartBuilder
        self.workbook = WorkbookHandler()
        self.chart_builder = ChartBuilder()

    def add_title_sheet(
        self,
        title: str,
        subtitle: str = "",
        metadata: Optional[Dict[str, Any]] = None
    ) -> Any:
        """
        Add a title/cover sheet to the report.

        Args:
            title: Report title.
            subtitle: Report subtitle.
            metadata: Dictionary of metadata key-value pairs.

        Returns:
            The created worksheet.
        """
        sheet = self.workbook.workbook.create_sheet("Cover", 0)

        # Title
        sheet['A1'] = title
        self.workbook.format_cells('Cover', 'A1:A1', bold=True, font_size=18)

        # Subtitle
        if subtitle:
            sheet['A2'] = subtitle
            self.workbook.format_cells('Cover', 'A2:A2', font_size=12)

        # Metadata
        if metadata:
            row = 4
            for key, value in metadata.items():
                sheet[f'A{row}'] = f"{key}:"
                sheet[f'B{row}'] = str(value)
                self.workbook.format_cells('Cover', f'A{row}:A{row}', bold=True)
                row += 1

        return sheet

    def add_data_sheet(
        self,
        sheet_name: str,
        data: List[List[Any]],
        headers: Optional[List[str]] = None,
        format_header: bool = True
    ) -> Any:
        """
        Add a data sheet with optional headers.

        Args:
            sheet_name: Name of the sheet.
            data: 2D list of data rows.
            headers: Optional list of column headers.
            format_header: Whether to format the header row.

        Returns:
            The created worksheet.
        """
        sheet = self.workbook.workbook.create_sheet(sheet_name)

        start_row = 1
        if headers:
            sheet.append(headers)
            if format_header:
                header_range = f"A1:{chr(64 + len(headers))}1"
                self.workbook.format_cells(
                    sheet_name, header_range,
                    bold=True, fill_color='4472C4'
                )
            start_row = 2

        for row_data in data:
            sheet.append(row_data)

        # Auto-size columns
        self.workbook.auto_size_columns(sheet_name)

        return sheet

    def add_summary_with_chart(
        self,
        sheet_name: str,
        data: Dict[str, float],
        chart_type: str = 'bar',
        chart_title: str = ""
    ) -> Any:
        """
        Add a summary sheet with a chart.

        Args:
            sheet_name: Name of the sheet.
            data: Dictionary of label-value pairs.
            chart_type: Type of chart ('bar', 'line', 'pie', etc.).
            chart_title: Title for the chart.

        Returns:
            The created worksheet.
        """
        sheet = self.workbook.workbook.create_sheet(sheet_name)

        # Write data
        row = 1
        for label, value in data.items():
            sheet[f'A{row}'] = label
            sheet[f'B{row}'] = value
            row += 1

        # Format
        self.workbook.format_cells(sheet_name, f'A1:A{row-1}', bold=True)

        # Add chart
        chart = self.chart_builder.create_chart(
            chart_type=chart_type,
            title=chart_title or f"{sheet_name} Summary"
        )

        # Configure chart data
        data_range = f'A1:B{row-1}'
        self.chart_builder.add_data_to_chart(
            chart, sheet, data_range,
            categories_col='A', values_col='B'
        )

        # Position chart
        sheet.add_chart(chart, 'D2')

        return sheet

    def save(self, output_path: str) -> None:
        """
        Save the report to a file.

        Args:
            output_path: Path to save the report.
        """
        self.workbook.save(output_path)
        logger.info(f"Report saved to: {output_path}")


def aggregate_data_from_files(
    file_pattern: str,
    aggregation_func: Callable[[str], Any]
) -> List[Any]:
    """
    Aggregate data from multiple files matching a pattern.

    Args:
        file_pattern: Glob pattern (e.g., "C:/Data/*.csv").
        aggregation_func: Function to process each file and return aggregated data.

    Returns:
        List of aggregated results from each file.
    """
    files = glob.glob(file_pattern)
    logger.info(f"Found {len(files)} files matching pattern: {file_pattern}")

    results = []
    for file_path in files:
        try:
            result = aggregation_func(file_path)
            results.append(result)
            logger.info(f"Processed: {os.path.basename(file_path)}")
        except Exception as e:
            logger.warning(f"Failed to process {file_path}: {e}")

    return results


def create_pivot_summary(
    data: List[Dict[str, Any]],
    group_by: str,
    sum_field: str
) -> Dict[str, Any]:
    """
    Create a pivot-style summary from data.

    Args:
        data: List of dictionaries containing data.
        group_by: Field name to group by.
        sum_field: Field name to sum.

    Returns:
        Dictionary of {group_value: total}.
    """
    summary: Dict[str, Any] = {}
    for row in data:
        key = row.get(group_by, 'Unknown')
        value = row.get(sum_field, 0)

        if key in summary:
            summary[key] += value
        else:
            summary[key] = value

    return summary
