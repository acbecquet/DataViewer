import pytest
import pandas as pd
import tkinter as tk
from unittest.mock import MagicMock
from processing import load_excel_file  # Modify based on your usage

@pytest.fixture
def sample_dataframe():
    """Return a simple pandas DataFrame for testing."""
    return pd.DataFrame({
        "Column1": [1, 2, 3],
        "Column2": ["A", "B", "C"]
    })

@pytest.fixture
def mock_tk_root():
    """Return a mocked Tkinter root window."""
    root = tk.Tk()
    yield root
    root.destroy()

@pytest.fixture
def mock_gui(mock_tk_root):
    """Return a mocked instance of the TestingGUI class."""
    from main_gui import TestingGUI
    return TestingGUI(mock_tk_root)

@pytest.fixture
def mock_file_manager(mock_gui):
    """Return a FileManager instance with mocked GUI dependencies."""
    from file_manager import FileManager
    return FileManager(mock_gui)

@pytest.fixture
def mock_plot_manager(mock_gui):
    """Return a PlotManager instance with mocked GUI dependencies."""
    from plot_manager import PlotManager
    return PlotManager(mock_gui)

@pytest.fixture
def sample_excel_file(tmp_path):
    """Create a temporary Excel file for testing."""
    file_path = tmp_path / "test_file.xlsx"
    df = pd.DataFrame({"Sample": [1, 2, 3], "Value": [4, 5, 6]})
    df.to_excel(file_path, index=False)
    return str(file_path)
