from setuptools import setup

setup(
    name="xtotxt",
    version="0.1.0",
    description="Command-line tool for extracting text from PDF, DOCX, and PPTX files",
    py_modules=["extract_text"],
    install_requires=[
        "pdfminer.six>=20221105",
        "python-docx>=0.8.11", 
        "python-pptx>=0.6.21",
    ],
    entry_points={
        "console_scripts": [
            "extract_text=extract_text:main",
        ],
    },
) 