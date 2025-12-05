"""
Integration tests for Excel Automation module.
"""

import pytest
from pathlib import Path

try:
    from modules.excel_automation import WorkbookHandler, ChartBuilder
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False


@pytest.mark.skipif(not EXCEL_AVAILABLE, reason="openpyxl not installed")
class TestWorkbookHandler:
    """Integration tests for WorkbookHandler."""

    @pytest.fixture
    def workbook(self) -> WorkbookHandler:
        """Create a new workbook for testing."""
        return WorkbookHandler()

    @pytest.fixture
    def temp_excel_file(self, temp_dir: Path) -> Path:
        """Create path for temporary Excel file."""
        return temp_dir / "test_workbook.xlsx"

    def test_create_new_workbook(self, workbook: WorkbookHandler):
        """Test creating a new workbook."""
        assert workbook.workbook is not None
        assert len(workbook.get_sheet_names()) == 1  # Default sheet

    def test_save_and_load_workbook(
        self,
        workbook: WorkbookHandler,
        temp_excel_file: Path
    ):
        """Test saving and loading a workbook."""
        # Save
        workbook.save(str(temp_excel_file))
        assert temp_excel_file.exists()

        # Load
        loaded = WorkbookHandler(str(temp_excel_file))
        assert loaded.workbook is not None

    def test_create_and_delete_sheet(self, workbook: WorkbookHandler):
        """Test creating and deleting sheets."""
        initial_count = len(workbook.get_sheet_names())

        # Create
        workbook.create_sheet("TestSheet")
        assert "TestSheet" in workbook.get_sheet_names()
        assert len(workbook.get_sheet_names()) == initial_count + 1

        # Delete
        workbook.delete_sheet("TestSheet")
        assert "TestSheet" not in workbook.get_sheet_names()

    def test_create_sheet_at_index(self, workbook: WorkbookHandler):
        """Test creating sheet at specific index."""
        workbook.create_sheet("FirstSheet", 0)

        assert workbook.get_sheet_names()[0] == "FirstSheet"

    def test_write_and_read_data(self, workbook: WorkbookHandler):
        """Test writing and reading data."""
        sheet_name = workbook.get_sheet_names()[0]

        # Write data
        data = [
            ["Name", "Age", "City"],
            ["Alice", 30, "New York"],
            ["Bob", 25, "Los Angeles"],
            ["Charlie", 35, "Chicago"]
        ]
        workbook.write_data(sheet_name, data)

        # Read data
        read_data = workbook.read_data(sheet_name, "A1:C4")

        assert read_data[0] == ["Name", "Age", "City"]
        assert read_data[1] == ["Alice", 30, "New York"]
        assert len(read_data) == 4

    def test_write_data_with_offset(self, workbook: WorkbookHandler):
        """Test writing data starting from different cell."""
        sheet_name = workbook.get_sheet_names()[0]

        data = [["Value1", "Value2"], ["Value3", "Value4"]]
        workbook.write_data(sheet_name, data, start_cell="B2")

        # Read back
        read_data = workbook.read_data(sheet_name, "B2:C3")

        assert read_data[0] == ["Value1", "Value2"]
        assert read_data[1] == ["Value3", "Value4"]

    def test_format_cells(
        self,
        workbook: WorkbookHandler,
        temp_excel_file: Path
    ):
        """Test cell formatting."""
        sheet_name = workbook.get_sheet_names()[0]

        # Write header
        workbook.write_data(sheet_name, [["Header1", "Header2"]])

        # Format as bold
        workbook.format_cells(sheet_name, "A1:B1", bold=True, font_size=14)

        # Save and reload to verify
        workbook.save(str(temp_excel_file))
        loaded = WorkbookHandler(str(temp_excel_file))
        sheet = loaded.get_sheet(sheet_name)

        assert sheet["A1"].font.bold is True
        assert sheet["A1"].font.size == 14

    def test_format_cells_with_fill(
        self,
        workbook: WorkbookHandler,
        temp_excel_file: Path
    ):
        """Test cell fill color."""
        sheet_name = workbook.get_sheet_names()[0]

        workbook.write_data(sheet_name, [["Colored"]])
        workbook.format_cells(sheet_name, "A1:A1", fill_color="FFFF00")

        workbook.save(str(temp_excel_file))
        loaded = WorkbookHandler(str(temp_excel_file))
        sheet = loaded.get_sheet(sheet_name)

        # Verify fill was applied (color format varies by openpyxl version)
        assert sheet["A1"].fill.fill_type == "solid"
        assert "FFFF00" in str(sheet["A1"].fill.start_color.rgb)

    def test_add_formula(self, workbook: WorkbookHandler):
        """Test adding formulas to cells."""
        sheet_name = workbook.get_sheet_names()[0]

        # Add numbers
        workbook.write_data(sheet_name, [[10], [20], [30]])

        # Add sum formula
        workbook.add_formula(sheet_name, "A4", "=SUM(A1:A3)")

        sheet = workbook.get_sheet(sheet_name)
        assert sheet["A4"].value == "=SUM(A1:A3)"

    def test_auto_size_columns(
        self,
        workbook: WorkbookHandler,
        temp_excel_file: Path
    ):
        """Test auto-sizing columns."""
        sheet_name = workbook.get_sheet_names()[0]

        # Write data with varying lengths
        data = [
            ["Short", "This is a much longer text"],
            ["X", "Y"]
        ]
        workbook.write_data(sheet_name, data)
        workbook.auto_size_columns(sheet_name)

        workbook.save(str(temp_excel_file))
        loaded = WorkbookHandler(str(temp_excel_file))
        sheet = loaded.get_sheet(sheet_name)

        # Column B should be wider than column A
        col_a_width = sheet.column_dimensions["A"].width
        col_b_width = sheet.column_dimensions["B"].width

        assert col_b_width > col_a_width

    def test_get_sheet_names(self, workbook: WorkbookHandler):
        """Test getting sheet names."""
        workbook.create_sheet("Sheet2")
        workbook.create_sheet("Sheet3")

        names = workbook.get_sheet_names()

        assert len(names) == 3
        assert "Sheet2" in names
        assert "Sheet3" in names

    def test_close_workbook(self, workbook: WorkbookHandler):
        """Test closing workbook."""
        workbook.close()

        assert workbook.workbook is None

    def test_load_nonexistent_file_raises(self, temp_dir: Path):
        """Test loading nonexistent file raises error."""
        with pytest.raises(Exception):
            handler = WorkbookHandler()
            handler.load(str(temp_dir / "nonexistent.xlsx"))

    def test_save_without_filepath_raises(self, workbook: WorkbookHandler):
        """Test saving without filepath raises error."""
        with pytest.raises(ValueError):
            workbook.save()

    def test_full_workflow(self, temp_excel_file: Path):
        """Test a complete Excel workflow."""
        # Create workbook
        wb = WorkbookHandler()

        # Create sheets
        wb.create_sheet("Sales")
        wb.create_sheet("Summary")

        # Delete default sheet
        default_sheet = wb.get_sheet_names()[0]
        if default_sheet != "Sales":
            wb.delete_sheet(default_sheet)

        # Add sales data
        sales_data = [
            ["Month", "Revenue", "Expenses"],
            ["January", 10000, 7500],
            ["February", 12000, 8000],
            ["March", 15000, 9000],
        ]
        wb.write_data("Sales", sales_data)

        # Format header
        wb.format_cells("Sales", "A1:C1", bold=True, fill_color="4472C4")

        # Add border to data
        wb.format_cells("Sales", "A1:C4", border=True)

        # Add summary with formula
        wb.write_data("Summary", [["Total Revenue"]])
        wb.add_formula("Summary", "B1", "=SUM(Sales!B2:B4)")

        # Auto-size
        wb.auto_size_columns("Sales")

        # Save
        wb.save(str(temp_excel_file))

        # Verify
        assert temp_excel_file.exists()

        # Reload and check
        loaded = WorkbookHandler(str(temp_excel_file))
        assert "Sales" in loaded.get_sheet_names()
        assert "Summary" in loaded.get_sheet_names()

        sales_sheet = loaded.get_sheet("Sales")
        assert sales_sheet["A1"].value == "Month"
        assert sales_sheet["B2"].value == 10000
