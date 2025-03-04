import timeit

setup = '''
from gui import TestingGUI
app = TestingGUI()
app.load_excel_file("test_data.xlsx")
'''

print("Report generation time:",
      timeit.timeit('app.generate_full_report()', setup=setup, number=10))