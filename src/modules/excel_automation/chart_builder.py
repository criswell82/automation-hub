"""Chart Builder for Excel automation."""

import logging
from typing import Optional

try:
    from openpyxl.chart import BarChart, LineChart, PieChart, Reference
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


class ChartBuilder:
    """Build charts in Excel workbooks."""

    def __init__(self):
        if not OPENPYXL_AVAILABLE:
            raise RuntimeError("openpyxl is required")
        self.logger = logging.getLogger(__name__)

    def create_bar_chart(self, sheet, data_range: str, title: str = "Bar Chart"):
        """Create a bar chart."""
        chart = BarChart()
        chart.title = title
        data = Reference(sheet, range_string=data_range)
        chart.add_data(data, titles_from_data=True)
        return chart

    def create_line_chart(self, sheet, data_range: str, title: str = "Line Chart"):
        """Create a line chart."""
        chart = LineChart()
        chart.title = title
        data = Reference(sheet, range_string=data_range)
        chart.add_data(data, titles_from_data=True)
        return chart

    def create_pie_chart(self, sheet, data_range: str, title: str = "Pie Chart"):
        """Create a pie chart."""
        chart = PieChart()
        chart.title = title
        data = Reference(sheet, range_string=data_range)
        chart.add_data(data, titles_from_data=True)
        return chart
