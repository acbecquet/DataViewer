import pytest
from gui import TestingGUI

@pytest.fixture
def app():
    root = tk.Tk()
    return TestingGUI(root)

def test_file_load(benchmark, app):
    benchmark(app.load_excel_file, "C:\Users\Alexander Becquet\Documents\Python\Python\TPM Data Processing Python Scripts\Standardized Testing GUI\T58G 510 Standard FOR TESTING.xlsx")

def test_report_generation(benchmark, app):
    benchmark(app.generate_full_report)
