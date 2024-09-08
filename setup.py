from setuptools import setup, find_packages

VERSION = "0.1.1"

LONG_DESCRIPTION = (
    "A Python library to export Pandas dataframes to "
    "pretty Excel workbooks of tables,"
    "with accent on 1) maximum simplicity and minimal effort"
    "(sensible defaults). "
    "2) some flexibility with fonts, number formats, header/tab colors, "
    "etc."
)


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
    url='https://github.com/yourusername/foo',  # Replace with your project's URL
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.6',
)
