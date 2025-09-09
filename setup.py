import os
from setuptools import setup, find_packages

def read_requirements():
    try:
        with open("requirements.txt", "r") as f:
            return [line.strip() for line in f if line.strip() and not line.startswith('#')]
    except FileNotFoundError:
        # Fallback to minimal requirements
        return [
            "matplotlib>=3.9.0",
            "numpy>=2.0.0", 
            "pandas>=2.2.0",
            "openpyxl>=3.1.0",
            "pillow>=10.4.0",
            "psutil>=6.0.0",
            "sqlalchemy>=2.0.0",
            "requests>=2.25.0",
            "opencv-python>=4.5.0",
            "python-pptx>=0.6.0",
            "XlsxWriter>=3.2.0",
            "tkintertable>=1.3.0",
            "python-dateutil>=2.9.0",
            "pytz>=2024.1",
            "packaging>=20.0",
            "scikit-learn>=1.3.0"
        ]

setup(
    name="standardized-testing-gui",
    version="0.0.4", 
    author="Charlie Becquet",
    description="Standardized Testing GUI for Data Analysis",
    packages=find_packages(),
    py_modules=[
        'main', 'main_gui', 'file_manager', 'plot_manager', 
        'report_generator', 'trend_analysis_gui', 'progress_dialog',
        'image_loader', 'viscosity_calculator', 'utils', 'processing',
        'database_manager', 'data_collection_window', 'test_selection_dialog',
        'test_start_menu', 'header_data_dialog'
    ],
    install_requires=read_requirements(),
    entry_points={
        'console_scripts': [
            'testing-gui=main:main',
        ],
    },
    include_package_data=True,
    python_requires=">=3.8",
)
