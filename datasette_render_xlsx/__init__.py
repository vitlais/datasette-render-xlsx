from datasette import hookimpl
import xlsxwriter
from io import BytesIO
from datasette.utils.asgi import Response

# The function renders max 100 rows. This can be overridden in the Metadata settings.
# It's never possible to render more rows than the setting max_returned_rows
MAX_ROWS = '100'


@hookimpl
def register_output_renderer(datasette):
    return{
        "extension": "xlsx",
        "render": render_xlsx,
    }


async def render_xlsx(datasette, database, table, columns, request):
    # Get max rendered rows from Metadata
    # TODO: Handle the event of more rows than max_returned_rows
    plugin_config = datasette.plugin_config(
        "datasette-render-xlsx", database=database, table=table
    )
    try:
        max_rows = plugin_config['max_rows']
    except KeyError:
        max_rows = MAX_ROWS
    except TypeError:
        max_rows = MAX_ROWS

    # Get JSON from Datasette
    path = f"{request.path[:-5]}.json?{request.query_string}&size={max_rows}"
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
        worksheet.write_row(row=row_number, col=0, data=row)
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
