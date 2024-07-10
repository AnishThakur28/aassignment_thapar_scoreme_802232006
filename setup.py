from setuptools import setup, find_packages

setup(
    name='pdf-to-excel-tool',
    version='1.0',
    packages=find_packages(),
    install_requires=[
        'pdfminer.six',
        'openpyxl',
    ],
    entry_points={
        'console_scripts': [
            'pdf-to-excel = pdf_to_excel.__init__:main'
        ],
    },
)
