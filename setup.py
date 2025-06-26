import os
from setuptools import setup, find_packages

# Read requirements
def read_requirements():
    with open("requirements.txt", "r") as f:
        return [line.strip() for line in f if line.strip() and not line.startswith('#')]

# Read long description
def read_long_description():
    if os.path.exists("README.md"):
        with open("README.md", "r", encoding="utf-8") as f:
            return f.read()
    return "Standardized Testing GUI for Data Analysis"

# Find all resource files
def find_resource_files():
    resource_files = []
    for root, dirs, files in os.walk("resources"):
        for file in files:
            resource_files.append(os.path.join(root, file))
    return resource_files

setup(
    name="standardized-testing-gui",
    version="3.0.0",
    author="Charlie Becquet",
    description="Standardized Testing GUI for Data Analysis",
    long_description=read_long_description(),
    long_description_content_type="text/markdown",
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
            'sdr-testing=main:main',  # Alternative command name
        ],
        'gui_scripts': [
            'testing-gui-windowed=main:main',  # Windows-specific windowed mode
        ],
    },
    include_package_data=True,
    package_data={
        '': ['resources/*', 'resources/**/*'],
    },
    data_files=[
        ('resources', find_resource_files()),
    ],
    python_requires=">=3.8",
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Science/Research",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Operating System :: OS Independent",
        "Topic :: Scientific/Engineering :: Information Analysis",
    ],
    keywords="testing data-analysis gui tkinter excel",
)