import os
import sys
import logging

from setuptools import setup, find_packages


logger = logging.getLogger(__name__)


version = "0.0.1"

# Please update tox.ini when modifying dependency version requirements
install_requires = [
    "coloredlogs",
    "requests",
    "configparser",
    "argparse",
    "xlsxwriter",
    "coloredlogs",
    "google-api-python-client",
    "google-auth-httplib2",
    "google-auth-oauthlib",
]

dev_extras = [
    "nose",
    "pep8",
    "tox",
    "aiounittest",
    "requests",
    "pympler",
    "pypandoc",
    "pandoc",
]

docs_extras = [
    "Sphinx>=1.0",  # autodoc_member_order = 'bysource', autodoc_default_flags
    "sphinx_rtd_theme",
    "sphinxcontrib-programoutput",
]

try:
    import pypandoc

    long_description = pypandoc.convert("README.md", "rst")
    long_description = long_description.replace("\r", "")

except (IOError, ImportError):
    import io

    with io.open("README.md", encoding="utf-8") as f:
        long_description = f.read()

setup(
    name="cecc2019",
    version=version,
    description="cecc2019",
    long_description=long_description,
    url="https://github.com/ph4r05/",
    author="Dusan Klinec",
    author_email="dusan.klinec@gmail.com",
    license="MIT",
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "Programming Language :: Python",
        "Programming Language :: Python :: 3.5",
        "Programming Language :: Python :: 3.6",
        "Topic :: Security",
    ],
    packages=find_packages(),
    include_package_data=True,
    python_requires=">=3.5",
    install_requires=install_requires,
    extras_require={
        "dev": dev_extras,
        "docs": docs_extras,
    },

    entry_points={
        'console_scripts': [
            'cecc2019 = cecc2019.generator:main',
        ],
    }

)
