from datasette import hookimpl
import xlsxwriter
from io import BytesIO
from datasette.utils.asgi import Response
from urllib.parse import urlparse, parse_qs, urlencode

MAX_ROWS = 'max'


def get_url_maxrows(url, maxrows):
    """
    Modifies the URL that is used to get
     the JSON-file by adding a value for
     _size and setting  _labels to off,
     since labels can't be rendered in this version.
    """
    parsed_url = urlparse(url)
    query = parse_qs(parsed_url.query)
    query['_size'] = [maxrows]
    query['_labels'] = ['off']
    new_query = urlencode(query, doseq=True)
    return parsed_url._replace(query=new_query).geturl()


@hookimpl
def register_output_renderer(datasette):

    return{
        "extension": "xlsx",
        "render": render_xlsx,
        "can_render": can_render_xslx,
    }


async def render_xlsx(datasette, database, table, columns, request):
    """
    This function gets the current data as JSON,
    writes data to xlsx and return the xlsx-file as response
    """
    # Get max rendered rows from Metadata settings
    plugin_config = datasette.plugin_config(
        "datasette-render-xlsx", database=database, table=table
    )
    try:
        max_rows = plugin_config['max_rows']
        if isinstance(max_rows, int) or max_rows.isnumeric:
            if int(max_rows) > datasette.setting('max_returned_rows'):
                max_rows = 'max'

    except KeyError:
        max_rows = MAX_ROWS
    except TypeError:
        max_rows = MAX_ROWS

    # Get JSON from Datasette
    path = f"{request.path[:-5]}.json?{request.query_string}"
    path = get_url_maxrows(path, max_rows)
    json_response = await datasette.client.get(path)
    json_object = json_response.json()

    rows = json_object['rows']
    # rows_count = len(rows)

    # Set output for xlsxwriter
    output = BytesIO()

    # Prepare workbook and worksheet
    workbook = xlsxwriter.Workbook(output,
                                   {
                                       'constant_memory': True,
                                       'in_memory': True,
                                       'strings_to_numbers': True,
                                   })
    bold = workbook.add_format({'bold': True})

    if table:
        sheet_name = table
    else:
        sheet_name = database
    worksheet = workbook.add_worksheet(sheet_name)

    # Write the first row with titles
    # For the future: It should be possible to use column.title from Metadata
    worksheet.write_row(row=0, col=0, data=columns, cell_format=bold)

    # Loop through JSON and write rows
    row_number = 1
    for row in rows:
        try:
            worksheet.write_row(row=row_number, col=0, data=row)
        except TypeError:
            workbook.close()
            return None
        else:
            row_number += 1

    worksheet.autofit()
    workbook.close()
    output.seek(0)

    # Prepare and return response
    response = Response(output.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    if table:
        workbook_name = f"{database} - {table}.xlsx"
    else:
        workbook_name = f"{database}.xlsx"
    response.headers = {'Content-Disposition': f'attachment; filename="{workbook_name}"'}

    return response


async def can_render_xslx(datasette, database, table, request):
    """
    This function returns false if the current data can't
    be rendered. It looks for two things:
    1. Explicit "do-not-render setting"
    2. Errors when reading JSON or writing xlsx
    """

    # Check if do_not_render is True
    plugin_config = datasette.plugin_config(
        "datasette-render-xlsx", database=database, table=table
    )
    try:
        if plugin_config['do_not_render']:
            return False
    except KeyError:
        pass
    except TypeError:
        pass

    # Get 100 rows of data and check if its readable
    path = f"{request.path}.json?{request.query_string}"
    path = get_url_maxrows(path, '100')
    json_response = await datasette.client.get(path)
    json_object = json_response.json()

    try:
        rows = json_object['rows']
    except KeyError:
        return False

    # Check if data is writable to xlsx-file
    # Set output for xlsxwriter
    output = BytesIO()

    # Prepare workbook and worksheet
    workbook = xlsxwriter.Workbook(output,
                                   {
                                       'constant_memory': True,
                                       'in_memory': True,
                                       'strings_to_numbers': True,
                                   })
    worksheet = workbook.add_worksheet()

    # Loop through JSON and write rows
    row_number = 1
    for row in rows:
        try:
            worksheet.write_row(row=row_number, col=0, data=row)
        except TypeError:
            workbook.close()
            return False
        else:
            row_number += 1
    workbook.close()
    return True
