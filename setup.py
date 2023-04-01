from setuptools import setup
import os

VERSION = "0.1"


def get_long_description():
    with open(
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "README.md"),
        encoding="utf8",
    ) as fp:
        return fp.read()


setup(
    name="datasette-render-xlsx",
    description="Download tables as xlsx-file",
    long_description=get_long_description(),
    long_description_content_type="text/markdown",
    author="Henrik Vitalis",
    url="https://github.com/vitlais/datasette-render-xlsx",
    project_urls={
        "Issues": "https://github.com/vitlais/datasette-render-xlsx/issues",
        "CI": "https://github.com/vitlais/datasette-render-xlsx/actions",
        "Changelog": "https://github.com/vitlais/datasette-render-xlsx/releases",
    },
    license="Apache License, Version 2.0",
    classifiers=[
        "Framework :: Datasette",
        "License :: OSI Approved :: Apache Software License"
    ],
    version=0.1,
    packages=["datasette_render_xlsx"],
    entry_points={"datasette": ["render_xlsx = datasette_render_xlsx"]},
    install_requires=["datasette", "xlsxwriter"],
    extras_require={"test": ["pytest", "pytest-asyncio"]},
    python_requires=">=3.7",
)
