import setuptools
from setuptools import Extension
import os

version = os.environ.get('RELEASE_VERSION', None)
if version is None:
    raise ValueError(f"version {version} is None. ENV var RELEASE_VERSION: {os.environ.get('RELEASE_VERSION')}")
version = version.lstrip('v')
print(f"verion: {version}")

setuptools.setup(
    name="champion-excel-form-extractor",
    version=version,
    url="https://github.com/adhadse/excelFormExtractor",
    author="Anurag Dhadse",
    description="Extract excel form content into structured data.",
    long_description_content_type="text/markdown",
    packages=setuptools.find_packages(include=["py_excel_form_extractor"]),
    license="MIT",
    platforms="Linux",
    keywords=["go", "golang", "python", "excel", "xlsx", "form", "extractor"],
    ext_modules=[
        Extension(
            "py_excel_form_extractor._extractor",
            sources=["py_excel_form_extractor/extractor.c"],
            include_dirs=["py_excel_form_extractor"],
        )
    ],
    py_modules = ["py_excel_form_extractor.extractor", "py_excel_form_extractor.utils"],
    package_data={"py_excel_form_extractor": [
        "*.so",
        "*_go.py",
        "*.py",
        "_*.py",
        "*.h",
        "*.c",
    ]},
    include_package_data=True,
)
