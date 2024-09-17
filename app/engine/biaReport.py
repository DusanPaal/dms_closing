# pylint: disable = C0103, E0110, E1101, R0912

"""
The 'biaReport.py' module provides following functionalities:
    - creating excel reports containing combined FBL5N and DMS data with processing output
    - summarizing processing output and selected data params into a HTML table row
    - uploading excel reports to a network location

Version history:
1.0.20210722 - initial version
1.0.20220211 - date-containing fields are converted from datetime to excel serial date.
               An explicit date format is then applied to the fields to display the vals
               correctly in reports.
"""

from datetime import datetime, date
from os.path import join
import pandas as pd
from pandas import ExcelWriter, DataFrame, Series
from xlsxwriter.workbook import Workbook
from xlsxwriter.format import Format

def _get_col_width(vals: Series, fld_name: str) -> int:
    """
    Returns excel column width calculated as
    the maximum count of characters contained
    in the column name and column data strings.
    """

    vals = vals.astype("string").dropna().str.len()
    vals = list(vals)
    vals.append(len(str(fld_name)))

    return max(vals)

def _col_to_rng(data: DataFrame, first_col: str, last_col: str = None,
                row: int = -1, last_row: int = -1) -> str:
    """
    Generates excel data range notation (e.g. 'A1:D1', 'B2:G2'). \n
    If 'last_col' is None, then only single-column range will be \n
    generated (e.g. 'A:A', 'B1:B1'). if 'row' is '-1', then the generated \n
    range will span through all column(s) rows (e.g. 'A:A', 'E:E').

    Params:
    -------
    data:
        Data for which colum names should be converted to a range.

    first_col:
        Name of the first column.

    last_col:
        Name of the last column.

    row:
        Index of the row for which the range will be generated.

    last_row:
        Index of the last row.

    Returns:
    --------
    Excel data range notation.
    """

    if isinstance(first_col, str):
        first_col_idx = data.columns.get_loc(first_col)
    elif isinstance(first_col, int):
        first_col_idx = first_col
    else:
        assert False, "Argument 'first_col' has invalid type!"

    first_col_idx += 1
    prim_lett_idx = first_col_idx // 26
    sec_lett_idx = first_col_idx % 26

    lett_a = chr(ord('@') + prim_lett_idx) if prim_lett_idx != 0 else ""
    lett_b = chr(ord('@') + sec_lett_idx) if sec_lett_idx != 0 else ""
    lett = "".join([lett_a, lett_b])

    if last_col is None:
        last_lett = lett
    else:

        if isinstance(last_col, str):
            last_col_idx = data.columns.get_loc(last_col)
        elif isinstance(last_col, int):
            last_col_idx = last_col
        else:
            assert False, "Argument 'last_col' has invalid type!"

        last_col_idx += 1
        prim_lett_idx = last_col_idx // 26
        sec_lett_idx = last_col_idx % 26

        lett_a = chr(ord('@') + prim_lett_idx) if prim_lett_idx != 0 else ""
        lett_b = chr(ord('@') + sec_lett_idx) if sec_lett_idx != 0 else ""
        last_lett = "".join([lett_a, lett_b])

    if row == -1:
        rng = ":".join([lett, last_lett])
    elif first_col == last_col and row != -1 and last_row == -1:
        rng = f"{lett}{row}"
    elif first_col == last_col and row != -1 and last_row != -1:
        rng = ":".join([f"{lett}{row}", f"{lett}{last_row}"])
    elif first_col != last_col and row != -1 and last_row == -1:
        rng = ":".join([f"{lett}{row}", f"{last_lett}{row}"])
    elif first_col != last_col and row != -1 and last_row != -1:
        rng = ":".join([f"{lett}{row}", f"{last_lett}{last_row}"])
    else:
        assert False, "Undefined argument combination!"

    return rng

def _write_data(data: DataFrame, wrtr: ExcelWriter, sht_name: str):
    """
    Writes a DataFrame object to an excel file.
    """

    # format headers
    data.columns = data.columns.str.replace("_", " ", regex = False)

    # write date to report file
    data.to_excel(wrtr, index = False, sheet_name = sht_name)

    # replace spaces in column names back with underscores for a better
    # field manupulation in the code
    data.columns = data.columns.str.replace(" ", "_", regex = False)

def _generate_formats(report: Workbook) -> dict:

    formats = {}

    # define sheet custom data formats
    formats["money"] = report.add_format({"num_format": "#,##0.00", "align": "center"})
    formats["category"] = report.add_format({"num_format": "000", "align": "center"})
    formats["date"]  = report.add_format({"num_format": "mm.dd.yyyy", "align": "center"})
    formats["general"] = report.add_format({"align": "center"})
    formats["header"] = report.add_format({
        "align": "center",
        "bg_color": "#F06B00",
        "font_color": "white",
        "bold": True
    })

    return formats

def _format_header(sht, data: DataFrame, header_idx: int, fmt):
    """
    Formats the header of the report sheet.
    """

    # format report header
    sht.conditional_format(
        _col_to_rng(data, data.columns[0], data.columns[-1], row = header_idx),
        {"type": "no_errors", "format": fmt}
    )

    # freeze data header row and set autofiler on all fields
    sht.freeze_panes(header_idx, 0)

def _get_column_format(col_name: str, formats: dict) -> Format:
    """
    Returns appropriate data format
    for a given report column.
    """

    if col_name == "Disputed_Amount":
        fmt = formats["money"]
    elif col_name == "DC_Amount":
        fmt = formats["money"]
    elif col_name == "Category":
        fmt = formats["category"]
    elif col_name in ("Solved_On", "Created_On"):
        fmt = formats["date"]
    else:
        fmt = formats["general"]

    return fmt

def _to_excel_serial(day: date) -> int:
    """
    Converts a datetime object
    into excel-compatible
    date integer serial format.
    """

    delta = day - datetime(1899, 12, 30).date()
    days = delta.days

    return days

def _convert_data(data: DataFrame) -> DataFrame:
    """Converts data columns to specific data types."""

    result = data.copy()

    # convert datetime format to excel native serial date format
    # in order to format the date vals correctly in the report
    for col_name in ("Solved_On", "Created_On"):
        result[col_name] = result[col_name].apply(
            lambda x: _to_excel_serial(x) if not pd.isna(x) else x
        )

    result["Category"] = pd.to_numeric(result["Category"]).astype("UInt8")

    return result

def create_report(data: DataFrame, file_path: str, sheet_name: str, field_order: list):
    """
    Creates processing output report in xlsx file format from source data.

    Params:
    -------
    data:
        Data to write to an Excel file.

    file_path:
        Path to the Excel report file.

    sheet_name:
        Name of the data-containing report sheet.

    field_order:
        List of field names defining the order of columns in the report data layout.

    Returns:
    --------
    None.
    """

    if not file_path.endswith(".xlsx"):
        raise ValueError("Unsupported file format used!")

    if sheet_name == "":
        raise ValueError("Sheet name cannot be an empty string!")

    if data.empty:
        raise ValueError("Argument 'data' contains no records!")

    # reorder fields and convert data to appropriate data types
    data = data.reindex(columns = field_order)
    converted = _convert_data(data)

    with ExcelWriter(file_path, engine = "xlsxwriter") as wrtr:

        _write_data(converted, wrtr, sheet_name)

        # get report data sheet and generate column formats
        sht = wrtr.sheets[sheet_name]
        formats = _generate_formats(wrtr.book)

        for col_name in converted.columns:
            col_width = _get_col_width(converted[col_name], col_name) + 2
            col_rng = _col_to_rng(converted, col_name)
            col_fmt = _get_column_format(col_name, formats)
            sht.set_column(col_rng, col_width, col_fmt)  # apply new column params

        _format_header(sht, converted, header_idx = 1, fmt = formats["header"])

def summarize(data: DataFrame, cocd: str, country: str) -> str:
    """
    Summarizes clearing results into a table \n
    and places the table into the HTML summary text.

    Params:
    ------
    data:
        Analyzed data, from which the summary will be created.

    cocd:
        Company code, for which the data is summarized.

    country:
        Name of the country, for which the data is summarized.

    Returns:
    --------
    A HTML table row contianing summarized data parameters.
    """

    # all cases that were solved (from status 1 to status 2)
    solved_cnt = data.query(
        "New_Status == 2 and Changed == True and IsError == False"
    )["Case_ID"].nunique()

    # all cases that were closed (from status 1/2 to status 2)
    closed_cnt = data.query(
        "New_Status == 3 and Changed == True and IsError == False"
    )["Case_ID"].nunique()

    # number of all open credit notes processed
    total_open_cnt = data[data["Clearing_Document"].isna()].shape[0]

    # number of items skipped due to incorrect case parameter combination
    inconsistent_cnt = data[data["Inconsistent"]].shape[0]

    # number of cases the prapeters of which were modified while keeping their oroginal status
    modified_cnt = data[data["Modified"]].shape[0]

    # number of cases where warnings were raised
    warnings_cnt = data[data["Warnings"].notna()].shape[0]

    # number of cases unprocessed due to an error raised by DMS
    errors_cnt = data[data["IsError"]].shape[0]

    # number of credit notes without case ID
    no_id_doc_cnt = data[data["Case_ID"].isna()].shape[0]

    # create HTML row summarizing data
    tbl_row = f"""
        <tr>
            <td style="border: purple 2px solid; padding: 5px">{country}</td>
            <td style="border: purple 2px solid; padding: 5px">{cocd}</td>
            <td style="border: purple 2px solid; padding: 5px">{modified_cnt}</td>
            <td style="border: purple 2px solid; padding: 5px">{solved_cnt}</td>
            <td style="border: purple 2px solid; padding: 5px">{closed_cnt}</td>
            <td style="border: purple 2px solid; padding: 5px">{inconsistent_cnt}</td>
            <td style="border: purple 2px solid; padding: 5px">{total_open_cnt}</td>
            <td style="border: purple 2px solid; padding: 5px">{no_id_doc_cnt}</td>
            <td style="border: purple 2px solid; padding: 5px">{warnings_cnt}</td>
            <td style="border: purple 2px solid; padding: 5px">{errors_cnt}</td>
        </tr>
    """

    return tbl_row

def create_notification(notification_path: str, template_path: str,
                        net_dir: str, net_subdir: str, summary: str):
    """
    Creates email HTML body of user notification.

    Params:
    -------
    notification_path:
        Path to the file to which the resulting notification will be written.

    template_path:
        Path to the file from which the notification template will be read.

    net_dir:
        Path to the network folder where all clearing reports are uploaded.

    net_subdir:
        Name of the subfolder, to which the current reports will be uploaded.

    summary:
        Summarized processing results in HTML format.

    Returns:
    --------
    None.
    """

    with open(template_path, 'r', encoding = "utf-8") as stream:
        template = stream.read()

    user_report_dir = join(net_dir, net_subdir)
    notif = template.replace("$ReportPath$", user_report_dir)
    notif = notif.replace("<tr><td>$TblRows$</td></tr>", summary)

    with open(notification_path, 'w', encoding = "utf-8") as stream:
        stream.write(notif)
