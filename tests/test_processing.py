# tests/test_processing.py
import pytest
import pandas as pd
import numpy as np
import sys
import os
from unittest.mock import Mock, patch
import matplotlib.pyplot as plt
# Add the project root to Python path so tests can find processing.py
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from processing import (
    get_valid_plot_options,
    validate_sheet_data,
    process_plot_sheet,
    aggregate_sheet_metrics,
    is_valid_excel_file,
    get_y_data_for_plot_type,
    get_y_label_for_plot_type,
    plot_tpm_bar_chart,
    plot_all_samples,
    convert_legacy_file_using_template,
    is_sheet_empty_or_zero,
    extract_samples_from_old_file,
    plot_aggregate_trends,
    process_generic_sheet,
    process_device_life_test

)

@pytest.mark.parametrize("plot_type,expected_label", [
    ("TPM", "TPM (mg/puff)"),
    ("Draw Pressure", "Draw Pressure (kPa)"),
    ("Invalid", "TPM (mg/puff)")  # Test default
])
def test_plot_labels(plot_type, expected_label):
    assert get_y_label_for_plot_type(plot_type) == expected_label


def test_get_valid_plot_options_basic():
    # Create sample data with 12 columns
    data = pd.DataFrame(np.random.rand(4, 12), columns=range(12))
    data.iloc[3:, 8] = [np.nan, 2, 3]  # TPM column
    data.iloc[3:, 3] = [0, 0, 0]  # Draw Pressure column
    valid = get_valid_plot_options(["TPM", "Draw Pressure"], data)
    assert valid == ["TPM"]

def test_get_valid_plot_options_all_invalid():
    # Create data with 12 columns
    data = pd.DataFrame(np.zeros((4, 12)))
    data.iloc[3:, 8] = [0, 0, 0]  # Invalid TPM
    valid = get_valid_plot_options(["TPM"], data)
    assert valid == []


# tests/test_processing.py
from processing import validate_sheet_data

def test_validate_sheet_data_empty():
    assert validate_sheet_data(pd.DataFrame()) is False  # Empty DataFrame

def test_validate_sheet_data_missing_columns():
    data = pd.DataFrame({"A": [1], "B": [2]})
    assert validate_sheet_data(data, required_columns=["C"]) is False

def test_validate_sheet_data_valid():
    data = pd.DataFrame({"RequiredCol": [1, 2, 3]})
    assert validate_sheet_data(data, required_columns=["RequiredCol"], required_rows=2) is True

# tests/test_processing.py
from processing import process_plot_sheet

@pytest.fixture
def mock_plot_sheet_data():
    data = pd.DataFrame(np.random.rand(20, 12))
    data = data.rename(columns={5: "Sample Name"})  # Proper column name
    data.iloc[3:, 8] = np.random.rand(17) * 5
    return data

def test_process_plot_sheet_valid(mock_plot_sheet_data):
    processed, samples, full_data = process_plot_sheet(mock_plot_sheet_data)
    
    # Verify 1 sample was processed
    assert len(processed) == 1
    assert "Sample Name" in processed["Sample Name"].values
    assert "Sample_1_Puffs" in samples  # Sample arrays exist

def test_process_plot_sheet_empty():
    processed, samples, full_data = process_plot_sheet(pd.DataFrame())
    assert processed.empty and not samples and full_data.empty

# tests/test_processing.py
from processing import aggregate_sheet_metrics

@pytest.fixture
def mock_full_sample_data():
    data = pd.DataFrame(np.random.rand(20, 12))
    
    # Set valid weight data
    data.iloc[3:, 2] = [1, 2, 3] + [np.nan] * 14  # 3 valid rows
    
    # Set TPM data
    data.iloc[3:, 8] = [1, 2, 3] + [np.nan] * 14
    return data

def test_aggregate_sheet_metrics(mock_full_sample_data):
    df = aggregate_sheet_metrics(mock_full_sample_data)
    
    # Check metrics for the sample
    assert df["Average TPM"].iloc[0] == pytest.approx(2.0)  # (1+2+3)/3 = 2
    assert df["Total Puffs"].iloc[0] == 3  # 3 valid rows in "after weight"

# tests/test_processing.py
from processing import is_valid_excel_file

def test_is_valid_excel_file():
    assert is_valid_excel_file("data.xlsx") is True
    assert is_valid_excel_file("~$data.xlsx") is False  # Temp file
    assert is_valid_excel_file("data.csv") is False  # Not Excel

# tests/test_processing.py
def test_process_plot_sheet_invalid_rows():
    # Data with fewer rows than required (data_start_row=4 needs >=5 rows)
    data = pd.DataFrame(np.random.rand(3, 12))
    processed, _, _ = process_plot_sheet(data)
    assert processed.empty

def test_aggregate_metrics_no_samples():
    df = aggregate_sheet_metrics(pd.DataFrame())  # Empty input
    assert df.empty


def test_get_y_data_and_labels():
    sample_data = pd.DataFrame(np.random.rand(10, 12))
    
    # Test TPM data extraction
    tpm_data = get_y_data_for_plot_type(sample_data, "TPM")
    assert len(tpm_data) == 7  # Rows 3 onward (10-3=7)
    
    # Test label mapping
    assert get_y_label_for_plot_type("Resistance") == "Resistance (Ohms)"
    assert get_y_label_for_plot_type("Unknown") == "TPM (mg/puff)"  # Default

def test_plot_tpm_bar_chart():
    fig, ax = plt.subplots()
    data = pd.DataFrame(np.random.rand(20, 24), 
                       columns=[f"Sample_{i}" for i in range(24)])
    sample_names = plot_tpm_bar_chart(ax, data, 2, 12)
    assert len(sample_names) == 2

def test_plot_all_samples_line():
    # Create 12 columns per sample
    data = pd.DataFrame({
        **{i: [1, 2, 3] for i in range(12)},
        **{i+12: [4, 5, 6] for i in range(12)}
    })
    fig, names = plot_all_samples(data, 12, "TPM")
    assert len(names) > 0

@patch("processing.load_workbook")
@patch("processing.extract_samples_from_old_file")
def test_legacy_conversion(mock_extract, mock_wb):
    # Mock a legacy sample
    mock_wb.return_value.sheetnames = ["Intense Test"]  # Add sheetnames mock
    mock_ws = Mock()
    mock_ws.max_row=20
    mock_wb.return_value.__getitem__.return_value = mock_ws

    mock_extract.return_value = [{
        "sample_name": "TestCart",
        "puffs": pd.Series([1, 2, 3]),
        "tpm": pd.Series([4, 5, 6])
    }]
    
    # Mock workbook operations
    mock_ws = Mock()
    mock_wb.return_value.__getitem__.return_value = mock_ws
    
    result = convert_legacy_file_using_template("dummy.xlsx")
    assert not result.empty
    mock_ws.cell.assert_called()  # Verify meta_data writing

def test_is_sheet_empty():
    empty_data = pd.DataFrame(np.zeros((16, 6)))  # 16 rows of zeros
    assert is_sheet_empty_or_zero(empty_data) is True
    
    valid_data = empty_data.copy()
    valid_data.iloc[5, 2] = 1  # Add one non-zero value
    assert is_sheet_empty_or_zero(valid_data) is False

def test_sample_extraction():
    mock_df = pd.DataFrame({
        "A": ["Cart #:", "Value", np.nan, "Puffs", 1, 2],
        "B": ["Voltage:", "5V", np.nan, "TPM (mg/puff)", 4, 5]
    })
    
    with patch("processing.read_sheet_with_values", return_value=mock_df):
        samples = extract_samples_from_old_file("dummy.xlsx")
        assert len(samples) == 1
        assert samples[0]["sample_name"] == "Cart #:"
        assert list(samples[0]["puffs"]) == [1, 2]

def test_aggregate_plot():
    data = pd.DataFrame({
        "Sample Name": ["A", "B"],
        "Average TPM": [5, 10],
        "Total Puffs": [100, 200]
    })
    
    fig = plot_aggregate_trends(data)
    ax1, ax2 = fig.axes
    
    # Verify bar plot
    assert len(ax1.patches) == 2
    assert ax1.get_ylabel() == "Average TPM (mg/puff)"
    
    # Verify line plot
    assert ax2.get_ylabel() == "Total Puffs"
    assert len(ax2.lines) == 1

# tests/test_processing.py
def test_generic_processing():
    data = pd.DataFrame(np.random.rand(10, 5))
    processed, _, _ = process_generic_sheet(data)
    assert not processed.empty

def test_device_life_test_processing():
    data = pd.DataFrame(np.random.rand(20, 12))
    processed, samples, _ = process_device_life_test(data)
    assert "Sample_1_Puffs" in samples



