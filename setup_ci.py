# https://github.com/tuananh/py-event-ruler/blob/main/setup_ci.py
import json
import os
import subprocess
import sys
import re
import setuptools
from setuptools import Extension
from setuptools.command.build_ext import build_ext


version = os.environ.get('RELEASE_VERSION', None)
if version is None:
    raise ValueError(f"version {version} is None. ENV var RELEASE_VERSION: {os.environ.get('RELEASE_VERSION')}")
version = version.lstrip('v')
print(f"verion: {version}")


def normalize(name):  # https://peps.python.org/pep-0503/#normalized-names
    return re.sub(r"[-_.]+", "-", name).lower()


PACKAGE_PATH = "py_excel_form_extractor"
PACKAGE_NAME = PACKAGE_PATH.split("/")[-1]
PACKAGE_DISTRIBUTION_NAME = "champion-excel-form-extractor"

if sys.platform == "darwin":
    # PYTHON_BINARY_PATH is setting explicitly for 310 and 311, see build_wheel.yml
    # on macos PYTHON_BINARY_PATH must be python bin installed from python.org or from brew
    PYTHON_BINARY = os.getenv("PYTHON_BINARY_PATH", sys.executable)
    # if PYTHON_BINARY == sys.executable:
    #     subprocess.check_call([sys.executable, "-m", "pip", "install", "pybindgen"])
else:
    # linux & windows
    PYTHON_BINARY = sys.executable
    # subprocess.check_call([sys.executable, "-m", "pip", "install", "pybindgen"])


def _generate_path_with_gopath() -> str:
    go_path = subprocess.check_output(["go", "env", "GOPATH"]).decode("utf-8").strip()
    path_val = f'{os.getenv("PATH")}:{go_path}/bin'
    return path_val


class CustomBuildExt(build_ext):
    def build_extension(self, ext: Extension):
        bin_path = _generate_path_with_gopath()
        dynamic_link = "False" if sys.platform == "darwin" else "True"
        go_env = json.loads(
            subprocess.check_output(["go", "env", "-json"]).decode("utf-8").strip()
        )

        destination = (
            os.path.dirname(os.path.abspath(self.get_ext_fullpath(ext.name)))
            + f"/{PACKAGE_NAME}"
        )
        # destination = PACKAGE_NAME

        subprocess.check_call(
            [
                "gopy",
                "build",
                "-no-make",
                f"-dynamic-link={dynamic_link}",
                "-rename=True",
                "-output",
                destination,
                "-vm",
                PYTHON_BINARY,
                *ext.sources,
            ],
            env={"PATH": bin_path, **go_env, "CGO_LDFLAGS_ALLOW": ".*"},
        )

        # dirty hack to avoid "from pkg import pkg", remove if needed
        os.makedirs(destination, exist_ok=True)
        with open(f"{destination}/__init__.py", "w") as f:
            f.write("from . import *")


with open("README.md") as f:
    readme = f.read()

with open("LICENSE") as f:
    license = f.read()

setuptools.setup(
    name=normalize(PACKAGE_DISTRIBUTION_NAME),
    version=version,
    url="https://github.com/championeng/excelFormExtractor",
    author="Anurag Dhadse",
    author_email="hello@adhadse.com",
    description="Extract excel form content into structured data.",
    long_description=readme,
    long_description_content_type="text/markdown",
    license=license,
    keywords=["go", "golang", "python", "excel", "xlsx", "form", "extractor"],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        # "Operating System :: OS Independent",
    ],
    packages=setuptools.find_packages(),
    cmdclass={
        "build_ext": CustomBuildExt,
    },
    ext_modules=[
        Extension(
            name=PACKAGE_NAME,
            sources=[
                # PACKAGE_PATH,
                "./pkg/extractor",
                "./pkg/utils",
            ],
            # include_dirs=["py_excel_form_extractor"],
        )
    ],
    # py_modules = ["py_excel_form_extractor.extractor", "py_excel_form_extractor.utils"],
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
