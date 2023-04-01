# datasette-render-xlsx

[![PyPI](https://img.shields.io/pypi/v/datasette-render-xlsx.svg)](https://pypi.org/project/datasette-render-xlsx/)
[![Changelog](https://img.shields.io/github/v/release/vitlais/datasette-render-xlsx?include_prereleases&label=changelog)](https://github.com/vitlais/datasette-render-xlsx/releases)
[![Tests](https://github.com/vitlais/datasette-render-xlsx/workflows/Test/badge.svg)](https://github.com/vitlais/datasette-render-xlsx/actions?query=workflow%3ATest)
[![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)](https://github.com/vitlais/datasette-render-xlsx/blob/main/LICENSE)


Download Datasette table as xlsx-file using XlsxWriter

## Installation

Install this plugin in the same environment as Datasette.

    datasette install datasette-render-xlsx

## Usage

This plugin adds a link for downloading the current table as an xslx-file that can
be opened and edited in e.g. MS Excel or LibreOffice.

The amount of rows rendered is limited by the value __max_returned_rows__ in Datasette's settings,
which defaults to 1000 rows. It is possible to set a
lower limit by using the plugin settings in Metadata:



    {
    "databases": {
            "sf-trees": {
                "plugins": {
                    "datasette-render-xsls": {
                        "max_rows": "500"
                    }
                }
            }
        }
    }`


Tables containing columns with blobs can't yet be rendered.
Labels that are used to display value from another table (foreign keys)
can't be rendered. The actual value will be shown.

You can turn off xlsx-rendering from specific databases or tables in the Metadata settings:

    {
      "databases": {
        "sakila": {
          "tables": {
            "category": {
              "plugins": {
                "datasette-render-xlsx": {
                  "do_not_render": true
                }
              }
            }
          }
        }
      }
    }

In this case there will be no link for download.


## Development

To set up this plugin locally, first checkout the code. Then create a new virtual environment:

    cd datasette-render-xlsx
    python3 -m venv venv
    source venv/bin/activate

Now install the dependencies and test dependencies:

    pip install -e '.[test]'

To run the tests:

    pytest
