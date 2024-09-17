# pylint: disable = C0103, C0301, W0703, W1203

"""
The 'biaProcessor.py' module performs all data-associated operations
such as parsing, cleaning, conversion, evaluation and querying.

Version history:
1.0.20210918 - initial version
1.0.20211025 - corrected a bug in convert_dms_data() when text lines in the 'txt'
               variable were not matched properly by regex pattern
1.0.20211026 - corrected a bug in search_matches() when case amount match
               was not evaluated correctly for situations where base threshold equals 0.0
1.0.20211222 - corrected a bug in search_matches() when all credit memo request numbers
               were replaced with the same credit note number in status sales containing
               multiple credit memo request numbers. Such cases will now be excluded from processing.
"""

from collections import namedtuple
from io import StringIO
import logging
from os.path import isfile
import re
import pandas as pd
from pandas import DataFrame, Series

Record = namedtuple("ClosingRecord", [
    "CaseID",
    "StatusSales",
    "RootCause",
    "Status"]
)

_STATUS_OPEN = 1
_STATUS_SOLVED = 2
_STATUS_CLOSED = 3
_STATUS_DEVALUATED = 4

_RC_UNUSED = ""
_RC_CREDIT_NOTE_ISSUED = "L06"
_RC_PAYMENT_AGREEMENT = "L01"
_RC_DISPUTE_UNJUSTIFIED = "L00"
_RC_CHARGE_OFF = "L08"
_RC_BELOW_THRESHOLD = "L14"

_logger = logging.getLogger("master")

def _generate_status_sales(old_val: str, credit_note: int) -> str:
    """
    Returns new 'Status sales' value containing the credit note number.
    """

    PRECREDIT_NOTE_RX = r"0?(50)\d{7}" # credit note number
    new_val = ""

    if str(credit_note) in old_val:
        new_val = pd.NA
    elif re.search(PRECREDIT_NOTE_RX, old_val) is not None:
        new_val = re.sub(PRECREDIT_NOTE_RX, str(credit_note), old_val)
    else:
        new_val = " ".join([old_val, str(credit_note)]).strip()

    return new_val

def _generate_ci_params(closed_items: DataFrame) -> DataFrame:
    """
    Generates new parameters for a case in DMS where
    the credit note is open on the customer account.
    """

    copied = closed_items.copy()
    subset = copied.query("Inconsistent == False")

    if subset.empty:
        return copied

    # set initial message for items with 'Status' value other than 'Open'. Further updates
    # may be performed however, such as inserting credit note number to Status Sales or updating root cause code.
    subset.loc[(subset["Status"] == _STATUS_CLOSED), "Message"] = "Credit note cleared, case already closed."

    # match items that contain new status and have sum of amounts below threshold
    matched = (subset["Status"].isin([_STATUS_OPEN, _STATUS_SOLVED]))
    subset.loc[matched, "New_Status"] = _STATUS_CLOSED
    subset.loc[matched, "Message"] = "Credit note cleared, case closed."

    # update status sales on credit note number, or leave the original
    # text if status sales already contains the number
    has_credit_node = (subset["Contains_Credit_Note"])

    subset.loc[has_credit_node, "Message"] = subset.loc[has_credit_node, "Message"].apply(
        lambda x: " ".join([x, "Status sales unchanged."])
    )

    subset.loc[~has_credit_node, "New_Status_Sales"] = subset.loc[~has_credit_node].apply(
        lambda x: _generate_status_sales(x["Status_Sales"], x["Document_Number"]), axis = 1
    )

    subset.loc[~has_credit_node, "Message"] = subset.loc[~has_credit_node, "Message"].apply(
        lambda x: " ".join([x, "Status sales updated."])
    )

    # update root cause code where the existing value is other than L00, L01, L06 or L14
    expected_rtc = subset["Root_Cause"].isin((
        _RC_CREDIT_NOTE_ISSUED, _RC_PAYMENT_AGREEMENT,
        _RC_BELOW_THRESHOLD, _RC_DISPUTE_UNJUSTIFIED
    ))

    subset.loc[expected_rtc, "Message"] = subset.loc[expected_rtc, "Message"].apply(
        lambda x: " ".join([x, "Root cause unchanged."])
    )

    subset.loc[~expected_rtc, "New_Root_Cause"] = _RC_CREDIT_NOTE_ISSUED

    subset.loc[~expected_rtc, "Message"] = subset.loc[~expected_rtc, "Message"].apply(
        lambda x: " ".join([x, "Root cause changed to L06."])
    )

    # select items that will be processed in DMS
    subset.loc[subset.query(
        "New_Status.notna() or "
        "New_Root_Cause.notna() or "
        "New_Status_Sales.notna()"
    ).index, "Changed"] = True

    subset.loc[subset.query(
        "New_Status.isna() and "
        "(New_Root_Cause.notna() or "
        "New_Status_Sales.notna())"
    ).index, "Modified"] = True

    # place the changed data subset
    # back to the copy of original data
    copied.loc[subset.index] = subset

    return copied

def _generate_oi_params(open_items: DataFrame) -> DataFrame:
    """
    Generates new parameters for a case in DMS where
    the credit note is already closed on the customer account.
    """

    # process consistent data only
    copied = open_items.copy()
    subset = copied.query("Inconsistent == False and Case_ID.notna()").copy()

    if subset.empty:
        return subset

    # set initial message for items with 'Status' value other than 'Open'. Further updates
    # may be performed however, such as inserting credit note number to Status Sales or updating root cause code.
    subset.loc[(subset["Status"] == _STATUS_CLOSED), "Message"] = "Case already closed."
    subset.loc[(subset["Status"] == _STATUS_SOLVED), "Message"] = "Case already solved."

    # match items that contain new status and have sum of amounts below threshold
    open_matched = (subset["Status"] == _STATUS_OPEN) & (subset["Amount_Match"])
    subset.loc[open_matched, "New_Status"] = _STATUS_SOLVED
    subset.loc[open_matched, "Message"] = "Case solved."

    open_unmatched = (subset["Status"] == _STATUS_OPEN) & (~subset["Amount_Match"])
    subset.loc[open_unmatched, "Message"] = "Case unsolved. Reason: Credit note and dispute amounts off threshold."

    # update status sales on credit note number, or leave the original
    # text if status sales already contains the number
    has_credit_node = (subset["Contains_Credit_Note"])

    subset.loc[has_credit_node, "Message"] = subset.loc[has_credit_node, "Message"].apply(
        lambda x: " ".join([x, "Status sales unchanged."])
    )

    subset.loc[~has_credit_node, "New_Status_Sales"] = subset.loc[~has_credit_node].apply(
        lambda x: _generate_status_sales(x["Status_Sales"], x["Document_Number"]), axis = 1
    )

    subset.loc[~has_credit_node, "Message"] = subset.loc[~has_credit_node, "Message"].apply(
        lambda x: " ".join([x, "Status sales updated."])
    )

    # update root cause code where the root cause code is other than L06, L01 or L14
    expected_rtc = subset["Root_Cause"].isin((_RC_CREDIT_NOTE_ISSUED, _RC_PAYMENT_AGREEMENT, _RC_BELOW_THRESHOLD))

    subset.loc[expected_rtc, "Message"] = subset.loc[expected_rtc, "Message"].apply(
        lambda x: " ".join([x, "Root cause unchanged."])
    )

    subset.loc[~expected_rtc, "New_Root_Cause"] = _RC_CREDIT_NOTE_ISSUED

    subset.loc[~expected_rtc, "Message"] = subset.loc[~expected_rtc, "Message"].apply(
        lambda x: " ".join([x, "Root cause changed to L06."])
    )

    copied.loc[subset.index] = subset

    # select items that will be processed in DMS
    copied.loc[copied.query(
        "New_Status.notna() or "
        "New_Root_Cause.notna() or "
        "New_Status_Sales.notna()"
    ).index, "Changed"] = True

    copied.loc[copied.query(
        "New_Status.isna() and "
        "(New_Root_Cause.notna() or "
        "New_Status_Sales.notna())"
    ).index, "Modified"] = True

    return copied

def _parse_amounts(vals: Series) -> Series:
    """
    Converts string amounts in the SAP
    format to floating point literals.

    Params:
    -------
    vals:
        String amounts in the SAP format.

    Returns:
    --------
    A Series object containing parsed floats.
    """

    repl = vals.str.replace(".", "", regex = False).str.replace(",", ".", regex = False)
    repl = repl.mask(repl.str.endswith("-"), "-" + repl.str.rstrip("-"))
    conv = pd.to_numeric(repl).astype("float64")

    return conv

def _read_fbl5n_data(file_paths: list) -> list:
    """
    Reads raw textual data exported from FBL5N. \n
    Returns a list of file contents stored as strings.
    """

    if len(file_paths) == 0:
        raise ValueError("Argument 'file_paths' is empty!")

    texts = []

    for f_path in file_paths:
        with open(f_path, 'r', encoding = "utf-8") as stream:
            texts.append(stream.read())

    return texts

def _preprocess_fbl5n_data(text: str) -> str:
    """
    Performs preprocessing operations on \n
    raw textual data exported from FBL5N:
    - extracting lines containing account items
    - removing leading and trailing pipes from lines
    - removing double quotes from text
    """

    matches = re.findall(r"^\|\s*\d{9}.*\|$", text, re.M)
    del text

    lines = "\n".join(matches)
    del matches

    replaced = re.sub(r"^\|", "", lines, flags = re.M)
    del lines

    replaced = re.sub(r"\|$", "", replaced, flags = re.M)
    replaced = re.sub(r"\"", "", replaced, flags = re.M)

    return replaced

def _parse_fbl5n_data(preproc: str) -> DataFrame:
    """
    Parses the preprocessed FBL5N textual \n
    data into a DataFrame object.
    """

    parsed = pd.read_csv(StringIO(preproc),
        sep = "|",
        dtype = "string",
        names =  [
            "Document_Number",
            "DC_Amount",
            "Branch",
            "Tax",
            "Text",
            "Company_Code",
            "Document_Assignment",
            "Clearing_Document"
        ]

    )

    return parsed

def _clean_fbl5n_data(parsed: DataFrame) -> DataFrame:
    """
    Cleans the parsed FBL5N data by removing \n
    non-printable leading and trailing spaces \n
    and removing misleading asterisk chars from \n
    tax sybols.
    """

    # remove leading and trailing non-printable chars form data
    for col in parsed.columns:
        parsed[col] = parsed[col].str.strip()

    # replace non-standard tax with empty strings
    cleaned = parsed.assign(
        tax = parsed["Tax"].replace("**", "")
    )

    return cleaned

def _convert_fbl5n_data(cleaned: DataFrame) -> DataFrame:
    """
    Converts columns of the cleaned FBL5N data \n
    to appropriate data types.
    """

    converted = cleaned.copy()
    converted["DC_Amount"] = _parse_amounts(cleaned["DC_Amount"])
    converted["Branch"] = pd.to_numeric(cleaned["Branch"]).astype("UInt64")
    converted["Document_Number"] = cleaned["Document_Number"].astype("UInt64")
    converted["Tax"] = cleaned["Tax"].astype("category")
    converted["Company_Code"] = cleaned["Company_Code"].astype("category")
    converted["Clearing_Document"] = pd.to_numeric(cleaned["Clearing_Document"]).astype("UInt64")

    return converted


def _read_dms_data(file_path: str) -> str:
    """
    Reads raw textual data exported from DMS. \n
    Returns file content stored as a string.
    """

    with open(file_path, 'r', encoding = "utf-8") as stream:
        txt = stream.read()

    return txt

def _preprocess_dms_data(text: str) -> str:
    """
    Performs preprocessing operations on \n
    raw textual data exported from FBL5N:
    - extracting lines containing account items
    - removing leading and trailing pipes from lines
    """

    matches = re.findall(r"^\|\s*\d+\s*\|.*$", text, re.M)
    del text

    lines = "\n".join(matches)
    del matches

    replaced = re.sub(r"^\|", "", lines, flags = re.M)
    replaced = re.sub(r"\|$", "", replaced, flags = re.M)

    return replaced

def _parse_dms_data(preproc: str):
    """
    Parses the preprocessed DMS textual \n
    data into a DataFrame object.
    """

    parsed = pd.read_csv(StringIO(preproc),
        sep = "|",
        dtype = "string",
        names = [
            "Case_ID",
            "Head_Office",
            "Debitor",
            "External_Reference",
            "Title",
            "Disputed_Amount",
            "Status_Sales",
            "Assignment",
            "Status",
            "Created_On",
            "Status_Description",
            "Short_Description_of_Customer",
            "Coordinator",
            "Processor",
            "Category_Description",
            "Root_Cause",
            "Created_By",
            "Category",
            "Solved_On"
        ]
    )

    return parsed

def _convert_dms_data(cleaned: DataFrame) -> DataFrame:
    """
    Converts columns of the cleaned DMS data \n
    into appropriate data types.
    """

    converted = cleaned.copy()
    converted["Case_ID"] = pd.to_numeric(cleaned["Case_ID"]).astype("UInt64")
    converted["Created_On"] = pd.to_datetime(cleaned["Created_On"],dayfirst = True).dt.date
    converted["Solved_On"] = pd.to_datetime(cleaned["Solved_On"], dayfirst = True).dt.date
    converted["Head_Office"] = pd.to_numeric(cleaned["Head_Office"]).astype("UInt64")
    converted["Debitor"] = pd.to_numeric(cleaned["Debitor"]).astype("UInt64")
    converted["Disputed_Amount"] = _parse_amounts(cleaned["Disputed_Amount"])
    converted["Status"] = cleaned["Status"].astype("UInt8")
    converted["Root_Cause"] = cleaned["Root_Cause"].astype("category")
    converted["Category"] = cleaned["Category"].astype("category")

    return converted

def _clean_dms_data(parsed: DataFrame) -> DataFrame:
    """
    Cleans the parsed FBL5N data by:
    - removing non-printable leading and trailing spaces
    - replacing missing vals.
    """

    for col in parsed.columns:
        parsed[col] = parsed[col].str.strip()

    # replace missing amounts with zero and
    # missing root cause codes with the 'NIL' flag.
    cleaned = parsed.assign(
        Disputed_Amount = parsed["Disputed_Amount"].replace("", "0,00"),
        Root_Cause = parsed["Root_Cause"].fillna("NIL")
    )

    return cleaned

def convert_fbl5n_data(file_paths: list) -> DataFrame:
    """
    Converts plain FBL5N text data into a panel dataset.

    Params:
    -------
    file_paths:
        Paths to the text files with FBL5N data to parse.

    Returns:
    --------
    Parsed panel data in the form of DataFrame object
    on success, None on failure.
    """

    texts = _read_fbl5n_data(file_paths)
    preproc = _preprocess_fbl5n_data("".join(texts))
    parsed = _parse_fbl5n_data(preproc)

    if parsed.empty:
        return None

    cleaned = _clean_fbl5n_data(parsed)
    converted = _convert_fbl5n_data(cleaned)

    return converted

def convert_dms_data(file_path: str) -> DataFrame:
    """
    Converts plain DMS text data to a panel dataset.

    Params:
    -------
    file_path:
        Path to the text file with DMS data to parse.

    Returns:
    --------
    Parsed panel data in the form of DataFrame object
    on success, None on failure.
    """

    if not isfile(file_path):
        raise ValueError(f"Path to the data file not found: {file_path}")

    text = _read_dms_data(file_path)
    preproc = _preprocess_dms_data(text)
    parsed = _parse_dms_data(preproc)

    if parsed.empty:
        return None

    cleaned = _clean_dms_data(parsed)
    converted = _convert_dms_data(cleaned)

    return converted

def assign_country(data: DataFrame, mapper: dict) -> DataFrame:
    """
    Assigns country of origin to the provided data based on company codes.

    Params:
    -------
    data:
        Source data containing a field with company codes.

    mapper:
        A map of company codes to country names.

    Returns:
    --------
    Panel dataset where each record is mapped to the country it belongs to.
    """

    if len(mapper) == 0:
        raise ValueError("Argument 'mapper' contains no records!")

    assigned = data.assign(
        Country = pd.NA
    )

    for cmp_cd in mapper:
        idx = data[data["Company_Code"] == cmp_cd].index
        assigned.loc[idx, "Country"] = mapper[cmp_cd]

    return assigned

def extract_cases(data: DataFrame, case_patts: dict) -> DataFrame:
    """
    Extracts case ID values contained the strings of the 'Text' field.

    Params:
    -------
    data:
        Source data containing the 'Text' field.

    case_patts:
        A map of country-specific case ID numbering to regex matching pattern.

    Returns:
    --------
    Input data enriched by the extracted case IDs in a separate field.
    """

    if data.empty:
        raise ValueError("Argument 'data' has no records!")

    if len(case_patts) == 0:
        raise ValueError("Argument 'case_patts' has no data!")

    extracted = data.assign(
        Case_IDs = pd.NA,
        Case_ID = pd.NA
    )

    for cocd in case_patts:

        idx = extracted[extracted["Company_Code"] == cocd].index

        if idx.empty:
            continue

        rx_patt = fr"D[P]?\s*[-_/]?({case_patts[cocd]})"
        extracted.loc[idx, "Case_IDs"] = extracted.loc[idx, "Text"].str.findall(rx_patt)
        extracted.loc[idx, "Case_ID"] = extracted.loc[idx, "Case_IDs"].apply(
            lambda x: x[0] if len(x) == 1 else pd.NA
        )

    extracted.drop("Case_IDs", axis = 1, inplace = True)
    extracted["Case_ID"] = pd.to_numeric(extracted["Case_ID"]).astype("UInt64")

    return extracted

def compact_data(acc_data: DataFrame, disp_data: DataFrame) -> DataFrame:
    """
    Merges preprocessed FBL5N accounting data with preprocessed DMS data \n
    into a single panel dataset, with 'Case_ID' field used as the merging key.

    Params:
    -------
    acc_data:
        Preprocessed FBL5N accounting data.

    disp_data:
        Preprocessed DMS disputed data.

    Returns:
    --------
    Panel dataset representing merged accounting and disputed data.
    """

    if acc_data.empty:
        raise ValueError("Argument 'acc_data' contains no records!")

    DUMMY_CASE_ID = 0

    acc_data["Case_ID"].fillna(DUMMY_CASE_ID, inplace = True)
    compacted = pd.merge(acc_data, disp_data, how = "left", on = "Case_ID")
    compacted["Case_ID"].mask(compacted["Case_ID"] == DUMMY_CASE_ID, pd.NA, inplace = True)
    compacted["Status_Sales"].fillna("", inplace = True)
    compacted.sort_values(["Case_ID"], ascending = False, inplace = True)

    compacted = compacted.assign(
        Changed = False,
        Inconsistent = False,
        Modified = False,
        IsError = False,
        Threshold = pd.NA,
        Warnings = pd.NA,
        Message = "",
        New_Status = pd.NA,
        New_Root_Cause = pd.NA,
        New_Status_Sales = pd.NA
    )

    return compacted

def check_consistency(data: DataFrame) -> DataFrame:
    """
    Checks data for any deviations form expected parameters. \n
    If any inconsistency is found, a detailed description is \n
    recorded for each affected case.

    Params:
    -------
    data:
        Merged FBL5N and DMS data.

    Returns:
    --------
    Validated data with recorded inconsistencies found.
    """

    MAX_CHARS = 50

    StandardRootCauses = (
        _RC_UNUSED,
        _RC_DISPUTE_UNJUSTIFIED,
        _RC_PAYMENT_AGREEMENT,
        _RC_CREDIT_NOTE_ISSUED,
        _RC_CHARGE_OFF,
        _RC_BELOW_THRESHOLD
    )

    copied = data.copy()

    copied = copied.assign(
        Contains_Credit_Note = False
    )

    checked = copied.query("Case_ID.notna()").copy()
    missing_id = copied.query("Case_ID.isna()").copy()

    missing_id["Warnings"] = "Case ID missing!"

    # check if status sales already
    # contains credit note number
    checked["Contains_Credit_Note"] = checked.apply(
        lambda x: str(x["Document_Number"]) in x["Status_Sales"], axis = 1
    )

    # inform the user about any invalid combinations of parameters per case
    valid_comb_a = checked["Contains_Credit_Note"] & (checked["Root_Cause"] == _RC_CREDIT_NOTE_ISSUED) & checked["Status"].isin((1, 2, 3))
    valid_comb_b = checked["Contains_Credit_Note"] & (checked["Root_Cause"] == _RC_DISPUTE_UNJUSTIFIED) & checked["Status"].isin((1, 2))
    valid_comb_c = ~checked["Contains_Credit_Note"] & checked["Root_Cause"].isin((_RC_UNUSED, _RC_CREDIT_NOTE_ISSUED)) & checked["Status"].isin((1, 2, 3))
    valid_comb_d = ~checked["Contains_Credit_Note"] & (checked["Root_Cause"] == _RC_PAYMENT_AGREEMENT) & (checked["Status"] == 2)
    valid_comb_e = checked["Contains_Credit_Note"] & (checked["Root_Cause"] == _RC_PAYMENT_AGREEMENT) & checked["Status"].isin((2, 3))

    valid_comb = valid_comb_a | valid_comb_b | valid_comb_c | valid_comb_d | valid_comb_e
    checked.loc[~valid_comb, "Message"] = "Case skipped. Reason: Incorrect case parameter combination!"
    checked.loc[~valid_comb, "Inconsistent"] = True

    # check data for entries with devaluated Case IDs
    # whose params cannot be changed in DMS
    deval = checked.query(f"Status == {_STATUS_DEVALUATED}")
    checked.loc[deval.index, "Message"] = "Case skipped. Reason: Devaluated Case ID assigned!"
    checked.loc[deval.index, "Inconsistent"] = True

    # check data for entries for invalid Case IDs
    inv_id = checked.query("Case_ID.notna() and Status.isna()")
    checked.loc[inv_id.index, "Message"] = "Case skipped. Reason: Invalid Case ID!"
    checked.loc[inv_id.index, "Inconsistent"] = True

    exceed_chars = checked.query(f"New_Status_Sales.str.len() > {MAX_CHARS}")
    checked.loc[exceed_chars.index, "Message"] = "Case skipped. Reason: Maximum number of 50 characters in 'Status sales' exceeded!"
    checked.loc[exceed_chars.index, "Inconsistent"] = True

    # check data for entries for invalid Case IDs
    id_not_found = checked.query("Case_ID.isna()")
    checked.loc[id_not_found.index, "Warnings"] = "Case ID not found in text!"
    checked.loc[id_not_found.index, "Inconsistent"] = True

    # check data for entries where FBL5N debitor differs from DMS debitor
    unequal_accs = checked.query("Branch != Debitor")
    checked.loc[unequal_accs.index, "Warnings"] = "FBL5N and DMS debitors not equal!"
    # this is kind of an insignificant inconsistency which does not prevent hte case to be processed in DMS
    checked.loc[unequal_accs.index, "Inconsistent"] = False

    # check data for entries where other than standard root cause is used
    unexpected_rtc = checked.query(f"Root_Cause not in {StandardRootCauses} and Case_ID.notna()")
    checked.loc[unexpected_rtc.index, "Warnings"] = "Unexpected root cause used!"
    checked.loc[unexpected_rtc.index, "Inconsistent"] = True

    multi_precredits = checked["Status_Sales"].str.contains(r"501\d{6}.*?501\d{6}", regex = True)
    checked.loc[multi_precredits, "Message"] = "Case skipped. Reason: Status sales contains multiple 501* numbers!"
    checked.loc[multi_precredits, "Inconsistent"] = True

    result = pd.concat([checked, missing_id])

    return result


def search_matches(data: DataFrame, cntry_rules: dict) -> DataFrame:
    """
    Searches data for cases to process based on defined criteria.

    Params:
    -------
    data:
        Merged FBL5N and DMS data containing credit notes and

    cntry_rules:
        Country-specific processing rules (base threshold, tax thresholds, etc...)

    Returns:
    --------
    Evaluated data.
    """

    if data.empty:
        raise ValueError("Argument 'data' contains no records!")

    if len(cntry_rules) == 0:
        raise ValueError("Argument 'cntry_rules' has no records!")

    bas_thresh = cntry_rules["base_threshold"]
    tax_threshs = cntry_rules["tax_thresholds"]
    cocd = cntry_rules["company_code"]

    # applies to situations/countries, where there's no tolerance limit
    # for the difference between disputed and credit note amounts
    if bas_thresh == 0:
        bas_thresh += 0.01

    subset = data[data["Company_Code"] == cocd]

    if subset.empty:
        _logger.warning(f"Data contains no records for company code '{cocd}'!")
        return subset

    # select  datasubset for a particuler company code from the entire dataset
    open_items = subset.query("Clearing_Document.isna() and Case_ID.notna()").copy()
    closed_items = subset.query("Clearing_Document.notna() and Case_ID.notna()").copy()
    missing_id = subset.query("Case_ID.isna()").copy()

    # sum document amounts based on case IDs - this is needed if there are > 1 case ID per credit note
    open_items["DC_Amount_Sum"] = open_items.groupby("Case_ID")["DC_Amount"].transform("sum")

    # calculate threshold values based on credit note tax codes
    open_items["Threshold"] = open_items["Tax"].apply(
        lambda x: tax_threshs[x] if x in tax_threshs else bas_thresh
    ).astype("float")

    # sum disputed amounts with DC amounts and compare the result with
    # the previously calculated threshold value
    open_items = open_items.assign(
        Total_Sum = open_items["DC_Amount_Sum"] + open_items["Disputed_Amount"]
    )

    # identify amounts that are below threshold
    open_items = open_items.assign(
        Amount_Match = open_items["Total_Sum"].abs() < open_items["Threshold"]
    )

    # generate new case params based on the above precalculations
    updated_oi = _generate_oi_params(open_items)
    updated_ci = _generate_ci_params(closed_items)

    # copy all updates to the original company code subset
    result = pd.concat([updated_oi, updated_ci, missing_id])

    assert result.shape[0] == subset.shape[0], "Original data and evaluate data contain different number of rows!"

    return result

def create_closing_input(data: DataFrame) -> list:
    """
    Creates data input used for case processing in DMS. \n
    The input is created based on the match evaluation \n
    results contained in the merged FBL5N and DMS data.

    Params:
    -------
    data:
        Output of the data evaluation process.

    Returns:
    --------
    Data used as the input for subsequent case modification in DMS.
    """

    recs = []
    subset = data.query("Inconsistent == False and (Changed == True or Modified == True)")

    tot_count = subset.shape[0]

    if tot_count == 0:
        _logger.warning("Could not create DMS closing input. Reason: No data to modify found.")
        return None

    _logger.debug(f"Total cases to process: {tot_count}.")

    _logger.info("Compiling input data for case processing in DMS ...")
    for _ , row in subset.iterrows():
        recs.append(Record(
            CaseID = row["Case_ID"],
            Status = row["New_Status"] if not pd.isna(row["New_Status"]) else None,
            RootCause = row["New_Root_Cause"] if not pd.isna(row["New_Root_Cause"]) else None,
            StatusSales = row["New_Status_Sales"] if not pd.isna(row["New_Status_Sales"]) else None
        ))

    return recs

def read_pickle(file_path: str) -> DataFrame:
    """
    Reads content of a pickled file.

    Params:
    -------
    file_path:
        Path to the file to read.

    Returns:
    --------
    A DataFrame object contatinig the file data.
    """

    if not file_path.endswith(".pkl"):
        raise ValueError("Unsupported file format used!")

    data = pd.read_pickle(file_path)

    return data
