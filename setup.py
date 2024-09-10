from setuptools import setup, find_packages

VERSION = "0.3"

from pathlib import Path
THIS_DIRECTORY = Path(__file__).parent
LONG_DESCRIPTION = (THIS_DIRECTORY / "README.md").read_text()

setup(
    name='excel_tables',
    version=VERSION,
    packages=find_packages(),
    install_requires=[
        # classes and data structures
        'pandas', 
        # excel
        'openpyxl', 'xlsxwriter',
        # presentation 
        'rich', 'babel', 'webcolors'

    ],
    author='Fralau',
    author_email='fralau@bluewin.ch',
    description='Python library to quickly export pandas tables to pretty, sensible Excel workbooks.',
    long_description=LONG_DESCRIPTION,
    long_description_content_type='text/markdown',
    url='https://github.com/fralau/excel_tables',  # Replace with your project's URL
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.6',
)
