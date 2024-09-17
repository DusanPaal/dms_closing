# pylint: disable = C0103, C0302, E0611, R1711, W0703, W1203

"""
The 'biaController.py' module represents the main communication
channel that manages data and control flow between the connected
highly specialized modules.

Version history
1.0.20210728 - initial version
1.0.20211025 - fixed regression in process_data() when function returned an empty
               list instead of None in situations when no cases to process were found
1.0.20220315 - added initialize_logger() procedure, removed clear_log() procedure
"""

from datetime import datetime, date, timedelta
from glob import glob
import json
import logging
from logging import config
from os import mkdir, remove
from os.path import exists, isfile, join, split
from shutil import move
import sys
from typing import Union

import yaml
from pandas import DataFrame, concat
from win32com.client import CDispatch

from . import biaDMS as dms
from . import biaFBL5N as fbl5n
from . import biaMail as mail
from . import biaProcessor as proc
from . import biaReport as rep
from . import biaSAP as sap
from . import biaDebugger as debug

_logger = logging.getLogger("master")


def get_current_date(fmt: str = None) -> Union[str,date]:
    """
    Returns a formatted current date.

    Params:
    -------
    fmt:
        The string that controls
        the format of the date.

    Returns:
    --------
    A string or a datatime.date object representing the current date.
    """

    curr_date = datetime.now().date()

    if fmt is None:
        return curr_date

    if fmt == "":
        return str(curr_date)

    return curr_date.strftime(fmt)

def get_past_date(day: date, day_offset: int, fmt: str = None) -> Union[str,date]:
    """
    Returns a past date calculated as a day offset to a given date.

    Params:
    -------
    day:
        The reference date from which the past date will be calculated.

    day_offset:
        Number of days to subtract from the reference day.

    fmt:
        The string that controls the format of the calculated date. \n
        If None is used (default value) then no string formatting \n
        of the calculated date will be performed and a raw datatime.date \n
        object will be returned instead.

    Returns:
    --------
    A string or a datatime.date object representing the past date.
    """

    assert day_offset > 0, "A day offset cannot be positive when the new day is in past!"

    past_day = min(day - timedelta(day_offset), day)

    if fmt is not None:
        return past_day.strftime(fmt)

    return past_day

def initialize_logger(cfg_path: str, log_path: str, header: dict, debug: bool = False) -> bool:
    """
    Creates a new or clears an existing log file and prints the log header.

    Params:
    ---------
    cfg_path:
        Path to a file with logging configuration params.

    log_path:
        Path to the application log file.

    header:
        Log header strings represented by parameter
        name (key) and description (value).

    debug:
        Indicates whether debug-level messages should be logged (default False).

    Returns:
    --------
    True if logger initialization succeeds, False if it fails.
    """

    try:
        with open(cfg_path, 'r', encoding = "utf-8") as stream:
            content = stream.read()
        log_cfg = yaml.safe_load(content)
        config.dictConfig(log_cfg)
    except Exception as exc:
        print (str(exc))
        return False

    if debug:
        _logger.setLevel(logging.DEBUG)
    else:
        _logger.setLevel(logging.INFO)

    prev_file_handler = _logger.handlers.pop(1)
    new_file_handler = logging.FileHandler(log_path)
    new_file_handler.setFormatter(prev_file_handler.formatter)
    _logger.addHandler(new_file_handler)

    try: # create a new / clear an existing log file
        with open(log_path, 'w', encoding = "utf-8"):
            pass
    except Exception as exc:
        print(str(exc))
        return False

    # write log header
    for i, (key, val) in enumerate(header.items(), start = 1):
        line = f"{key}: {val}"
        if i == len(header):
            line += "\n"
        _logger.info(line)

    return True

def load_app_config(cfg_path: str, states_path: str) -> dict:
    """
    Reads application configuration parameters.

    Params:
    -------
    cfg_path:
        Path to a .yaml file containing
        the configuration params.

    states_path:
        Path to a .json file containing
        application runtime states.

    Returns:
    --------
    Application configuration parameters. \n
    If loading fails due to an error, then None is returned.
    """

    _logger.info("Loading application configuration ...")

    try:
        with open(cfg_path, 'r', encoding = "utf-8") as stream:
            txt = stream.read()
    except Exception as exc:
        _logger.critical(f"Failed to load application configuration. Reason: {exc}")
        return None

    txt = txt.replace("$appdir$", sys.path[0])
    cfg = yaml.safe_load(txt)
    cfg.update({"states": {}})

    _logger.info("Loading application runtime states ...")

    try:
        with open(states_path, 'r', encoding = "utf-8") as stream:
            states = json.loads(stream.read())
    except Exception as exc:
        _logger.critical(f"Failed to load application states. Reason: {exc}")
        return None

    cfg["states"].update({"params": {}})

    for state in states:

        if state == "last_run":
            value = datetime.strptime(states[state], "%Y-%m-%d").date()
        else:
            value = state # in later versions more states may be needed

        cfg["states"].update({state: value})

    return cfg

def load_closing_rules(file_path: str) -> dict:
    """
    Reads and parses a file that contains data \n
    processing parameters used for evalation \n
    and subsequent closing of cases in DMS.

    Params:
    -------
    file_path:
        Path to a .yaml file containing the processing rules.

    Returns:
    --------
    A dict that maps countries (keys) to their \n
    respective closing rules (values).
    """

    _logger.info("Loading closing rules ...")

    try:
        with open(file_path, encoding = "utf-8") as stream:
            rules = yaml.safe_load(stream.read())
    except Exception as exc:
        _logger.critical(str(exc))
        return None

    return rules

def get_active_countries(rules: dict) -> dict:
    """
    Returns active countries to process.

    Params:
    -------
    rules:
        Data processing rules for particular countries.

    Returns:
    --------
    A map of countries (keys) and their company codes (vals).
    """

    _logger.info("Listing countries to process ...")

    countries = {}

    for cntry in rules:

        if not rules[cntry]["active"]:
            _logger.warning(f"{cntry} excluded form processing as per settings in country rules.")
            continue

        cocd = rules[cntry]["company_code"]
        countries.update({cntry: cocd})

    if len(countries) == 0:
        _logger.critical("No active country found!")
        return None

    _logger.info(f"Number of countries to process: {len(countries)}.")
    _logger.debug(f"Countries: {'; '.join(countries.keys())}.")

    return countries

def connect_to_sap(sap_cfg: dict) -> CDispatch:
    """
    Manages connecting of the application \n
    to the SAP GUI scripting engine.

    Params:
    -------
    sap_cfg:
        Application 'sap' configuration params.

    Returns:
    --------
    A SAP GuiSession object. \n
    If the attempt to connect to the scripting engine \n
    fails due to an error, then None is returned.
    """

    if sap_cfg["system"] == "P25":
        system = sap.Systems.P25
    elif sap_cfg["system"] == "Q25":
        system = sap.Systems.Q25

    _logger.info("Logging to SAP ... ")
    _logger.debug(f"System: '{system}'")

    try:
        sess = sap.login(sap_cfg["gui_exe_path"], system)
    except sap.LoginError as exc:
        _logger.critical(str(exc))
        return None

    return sess

def disconnect_from_sap(sess: CDispatch):
    """
    Manages disconnecting from
    the SAP GUI scripting engine.

    Params:
    -------
    sess:
        A SAP GuiSession object.

    Returns:
    --------
    None.
    """

    sap.logout(sess)

    return sess

def save_states(file_path: str, new_vals: dict = None):
    """
    Saves application runtime processig states to a file.

    Params:
    -------
    file_path:
        Path to the file containing the processig states.

    new_vals:
        Parameter names and their values to store. \n
        If None is used, then only 'last_run' parameter \n
        and a current date as its value will be stored.

    Returns:
    --------
    None.
    """

    _logger.info("Saving runtime params ...")

    # save application run date
    if new_vals is None:
        new_vals = {"last_run": get_current_date(fmt = "")}

    with open(file_path, 'r', encoding = "utf-8") as stream:
        states = json.load(stream)

    for key in new_vals:
        states[key] = new_vals[key]

    with open(file_path, 'w', encoding = "utf-8") as stream:
        json.dump(states, stream, indent = 4)

def export_fbl5n_data(data_cfg: dict, sap_cfg: dict, stat_cfg: dict,
                      countries: dict, sess: CDispatch) -> bool:
    """
    Manages data export from customer accounts into a local file.

    Params:
    -------
    data_cfg:
        Application 'data' configuration parameters.

    sap_cfg:
        Application 'sap' configuration parameters.

    stat_cfg:
        Application 'states' configuration parameters.

    countries:
        List of countries for which data will be exported.

    sess:
        A SAP GuiSession object.

    Returns:
    --------
    True if data export succeeds, False if it fails.
    """

    _logger.info("Starting FBL5N ...")
    fbl5n.start(sess)

    exp_path = join(data_cfg["export_dir"], data_cfg["fbl5n_export_name"])
    clr_exp_path = exp_path.replace("$type$", "cleared")
    opn_exp_path = exp_path.replace("$type$", "open")

    if exists(clr_exp_path):
        _logger.warning("Data for cleared items already exported from FBL5N in the previous run.")
    else:
        _logger.info("Exporting cleared items from FBL5N ...")

        try:
            fbl5n.export(clr_exp_path,
                layout = sap_cfg["fbl5n_layout"],
                company_codes = countries.values(),
                from_clr_date = get_past_date(stat_cfg["last_run"], data_cfg["days_closed"]),
                to_clr_date = get_current_date()
            )
        except fbl5n.NoDataFoundWarning as wng:
            # In case of cleared items, there might
            # be no cleared for a given range of days.
            _logger.warning(wng)
        except Exception as exc:
            _logger.exception(exc)
            _logger.info("Closing FBL5N ...")
            fbl5n.close()
            return False

    if exists(opn_exp_path):
        _logger.warning("Data for open items already exported from FBL5N in the previous run.")
    else:
        _logger.info("Exporting open items from FBL5N ...")
        try:
            fbl5n.export(opn_exp_path,
                layout = sap_cfg["fbl5n_layout"],
                company_codes = countries.values()
            )
        except fbl5n.NoDataFoundWarning as exc:
            # In case of open items, there must always
            # be some open items found on accs. If not,
            # this is an error rather than a warning.
            _logger.error(str(exc))
            return False
        except Exception as exc:
            _logger.exception(exc)
            _logger.info("Closing FBL5N ...")
            fbl5n.close()
            return False

    _logger.info("Closing FBL5N ...")
    fbl5n.close()

    return True

def export_dms_data(data_cfg: dict, sap_cfg: str, fbl5n_data: DataFrame, sess: CDispatch) -> bool:
    """
    Manages data export from DMS into a local data file.

    Params:
    -------
    data_cfg:
        Application 'data' configuration parameters.

    sap_cfg:
        Application 'sap' configuration parameters.

    fbl5n_data:
        Converted FBL5N data.

    sess:
        A SAP GuiSession object.

    Returns:
    --------
    True if data export succeeds, False if it fails.
    """

    exp_path = join(data_cfg["export_dir"], data_cfg["dms_export_name"])

    if exists(exp_path):
        _logger.warning("DMS data already exported in the previous application run.")
        return True

    subset = fbl5n_data.query("Case_ID.notna()")
    cases = tuple(subset["Case_ID"].unique())
    n_total = len(cases)

    _logger.info("Starting DMS ...")
    srch_mask = dms.start(sess)
    grid_view, n_found = dms.search_disputes(srch_mask, cases)

    if grid_view is None:
        _logger.info("Closing DMS ...")
        dms.close()
        return False

    if n_found < n_total:
        _logger.warning(f"Incorrect disputes detected: {n_total - n_found}")

    _logger.info("Exporting DMS data ...")

    try:
        dms.export(grid_view, exp_path, sap_cfg["dms_layout"])
    except Exception as exc:
        _logger.error(str(exc))
        exported = False
    else:
        exported = True
    finally:
        _logger.info("Closing DMS ...")
        dms.close()

    return exported

def preprocess_fbl5n_data(data_cfg: dict, rules: dict, countries: dict) -> dict:
    """
    Manages preprocessing of exported FBL5N data.

    Params:
    -------
    data_cfg:
        Application 'data' configuration parameters.

    rules:
        Data processing rules for particular countries.

    countries:
        List of countries for which accounting data will be preprocessed.

    Returns:
    --------
    A dict of preprocessed FBL5N data per company code. \n
    If conversion fails due to an error, then None is returned.
    """

    clr_exp_name = data_cfg["fbl5n_export_name"].replace("$type$", "cleared")
    opn_exp_name = data_cfg["fbl5n_export_name"].replace("$type$", "open")
    clr_exp_path = join(data_cfg["export_dir"], clr_exp_name)
    opn_exp_path = join(data_cfg["export_dir"], opn_exp_name)

    exp_paths = []

    if exists(clr_exp_path):
        exp_paths.append(clr_exp_path)

    if exists(opn_exp_path):
        exp_paths.append(opn_exp_path)

    _logger.info("Converting exported FBL5N data ...")
    converted = proc.convert_fbl5n_data(exp_paths)

    if converted is None:
        return None

    cocd_to_rx = {}
    cocd_to_cntry = {}

    for cntry in countries:
        cocd = rules[cntry]["company_code"]
        cocd_to_rx[cocd] = rules[cntry]["case_rx"]
        cocd_to_cntry[cocd] = cntry

    _logger.info("Extracting case ID numbers from data ...")
    extracted = proc.extract_cases(converted, cocd_to_rx)

    _logger.info("Assigning countries to data ...")
    assigned = proc.assign_country(extracted, cocd_to_cntry)

    return assigned

def preprocess_dms_data(data_cfg: dict) -> dict:
    """
    Manages preprocessing of exported DMS data.

    Params:
    -------
    data_cfg:
        Application 'data' configuration parameters.

    Returns:
    --------
    A dict of preprocessed DMS data per company code. \n
    If conversion fails due to an error, then None is returned.
    """

    _logger.info("Converting exported DMS data ...")
    exp_path = join(data_cfg["export_dir"], data_cfg["dms_export_name"])
    preproc = proc.convert_dms_data(exp_path)

    return preproc

def process_data(fbl5n_data: DataFrame, dms_data: DataFrame, countries: list, rules: dict) -> tuple:
    """
    Merges, compacts and checks the resulting data consistency, \n
    then creates processing input data containing parameters \n
    required for case processing in DMS.

    Params:
    -------
    fbl5n_data:
        Preprocessed FBL5N data.

    dms_data:
        Preprocessed DMS data.

    countries:
        List of countries for which data will be preprocessed.

    rules:
        Data processing rules for particular countries.

    Returns:
    --------
    A tuple of generated DMS processing input and the original \n
    data with consistency check results added as a separate field.
    """

    _logger.info("Compacting DMS and FBL5N data ...")

    compacted = proc.compact_data(fbl5n_data, dms_data)
    checked = proc.check_consistency(compacted)
    search_results = []

    for cntry in countries:
        _logger.info(f"Searching cases to process for {cntry} ...")
        search_res = proc.search_matches(checked, rules[cntry])
        search_results.append(search_res)

    concatenated = concat(search_results)
    closing_input = proc.create_closing_input(concatenated)

    return (closing_input, concatenated)

def check_output(cfg_data: dict) -> DataFrame:
    """
    Checks whether a dumped processing output \n
    from a previous application run exists.

    Params:
    -------
    cfg_data:
        Application 'data' configuration params.

    Returns:
    --------
    A DataFrame object containing the dumped data
    if a dump file is found. \n
    If no dump file is found, then None is returned.
    """

    _logger.info("Searching for a dumped output ...")

    file_path = join(cfg_data["dump_dir"], cfg_data["output_name"])

    if not isfile(file_path):
        _logger.info("No previous memory dump found.")
        return None

    _logger.info("Loading the dumped output ...")
    data = proc.read_pickle(file_path)

    return data

def _get_new_status(prev_val: int) -> dms.CaseStates:
    """
    Converts case status value stored as an \n
    integer to a dms.CaseStates enumerated value.
    """

    if prev_val is None:
        new_val = dms.CaseStates.Original
    elif prev_val == 2:
        new_val = dms.CaseStates.Solved
    elif prev_val == 3:
        new_val = dms.CaseStates.Closed
    else:
        assert False, "The used status is not valid for DMS closing!"

    return new_val

def _get_new_root_cause(prev_val: str) -> dms.RootCauses:
    """
    Converts root cause value stored as a \n
    string to a dms.RootCauses enumerated value.
    """

    if prev_val is None:
        new_val = None
    elif prev_val == "L06":
        new_val = dms.RootCauses.CREDITNOTE_ISSUED
    elif prev_val == "L01":
        new_val = dms.RootCauses.PAYMENT_AGREMENT
    else:
        assert False, "The used root cause is not valid for DMS closing!"

    return new_val

def process_disputes(closing_input: list, compacted: DataFrame, sess: CDispatch) -> DataFrame:
    """
    Modifies DMS cases with new parameters, changes case state where applicable \n
    and writes the DMS processing output message into a copy of the original data \n
    for each processed case.

    Params:
    -------
    closing_input:
        Input data for DMS processing (closing).

    compacted:
        The original compacted data.

    sess:
        A SAP GuiSession object.

    Returns:
    --------
    A DataFrame object representing the result of case processing. \n
    If the processing fails due to an error, then None is returned.
    """

    output = compacted.copy()

    _logger.info("Starting DMS ...")
    srch_mask = dms.start(sess)

    if srch_mask is None:
        return None

    for counter, rec in enumerate(closing_input, start = 1):

        case_id = rec.CaseID
        new_root_cause = rec.RootCause
        new_status_sales = rec.StatusSales

        _logger.info(f"Processing case: {case_id} ({counter} of {len(closing_input)}) ...")
        new_status = _get_new_status(rec.Status)
        new_root_cause = _get_new_root_cause(rec.RootCause)

        try:
            grid_view = dms.search_dispute(srch_mask, case_id)
        except Exception as exc:
            output.loc[output[output["Case_ID"] == case_id].index, "IsError"] = True
            output.loc[output["Case_ID"] == case_id, "Message"] = f"Case unprocessed. Error: {exc}"
            _logger.error(f" Processing failed. Reason: {exc}")
            continue

        try:
            dms.modify_case_parameters(grid_view, new_root_cause, new_status_sales, new_status)
        except Exception as exc:
            output.loc[output[output["Case_ID"] == case_id].index, "IsError"] = True
            output.loc[output["Case_ID"] == case_id, "Message"] = f"Case unprocessed. Error: {exc}"
            _logger.error(f" Case unprocessed. Error: {exc}")
            continue

    _logger.info("Closing DMS ...")
    dms.close()

    return output

def _create_reports(data: DataFrame, report_cfg: dict, notif_cfg: dict, countries: dict) -> bool:
    """
    Manages creation of user reports.

    Params:
    -------
    data:
        Data that will be written to the Excel report.

    report_cfg:
        Application 'reports' configuration params.

    notif_cfg:
        Application 'notifications' configuration params.

    countries:
        Countries and their company codes for which the reports will be created.

    Returns:
    --------
    True, if the creation of reports succeeds, False if it fails.
    """

    _logger.info("Creating user reports ...")

    summ = ""
    success = False

    for cntry in countries:

        _logger.info(f" Creating report and processing summary for {cntry} ...")

        cocd = countries[cntry]
        subset = data[data["Company_Code"] == cocd]
        file_name = report_cfg["report_name"].replace("$country$", cntry)
        file_name = file_name.replace("$company_code$", cocd)
        loc_file_path = join(report_cfg["local_report_dir"], file_name)

        try:
            rep.create_report(subset,
                loc_file_path,
                report_cfg["sheet_name"],
                report_cfg["field_order"]
            )
        except Exception as exc:
            _logger.error("Could not create report!", exc_info = exc)
            continue

        tbl_row = rep.summarize(subset, cocd, cntry)
        summ = "".join([summ, tbl_row]) # add the row to the existing summary
        success = True

    if not success:
        return False

    summary_path = join(notif_cfg["notification_dir"], notif_cfg["summary_name"])
    assert summ != "", "Summary is empty!"

    try:
        with open(summary_path, 'w', encoding = "utf-8") as stream:
            stream.write(summ)
    except Exception as exc:
        _logger.exception(exc)
        return False

    return True

def _create_notification(notif_cfg: dict, report_cfg: dict) -> bool:
    """
    Manages creation of user notifications.

    Params:
    -------
    notif_cfg:
        Application 'notifications' configuration params.

    data_cfg:
        Application 'data' configuration params.

    report_cfg:
        Application 'reports' configuration params.

    Returns:
    --------
    True, if the notification creation succeeds, False if it fails.
    """

    _logger.info("Creating user notification ...")

    summary_path = join(notif_cfg["notification_dir"], notif_cfg["summary_name"])

    try:
        with open(summary_path, 'r', encoding = "utf-8") as stream:
            summ = stream.read()
    except Exception as exc:
        _logger.exception(exc)
        return False

    try:
        rep.create_notification(
            notification_path = join(notif_cfg["notification_dir"], notif_cfg["notification_name"]),
            template_path = notif_cfg["template_path"],
            net_dir = report_cfg["net_report_dir"],
            net_subdir = get_current_date(report_cfg["net_report_subdir_format"]),
            summary = summ
        )
    except Exception as exc:
        _logger.exception(exc)
        return False

    return True

def _upload_reports(src_dir: str, dst_dir: str, dst_subdir: str) -> bool:
    """
    Manages uploading of user reports to a network folder.

    Creates a new subfolder in the destination folder (if this does not already exist) \n
    and moves the excel report file from a local source folder to the subfolder.

    Params:
    -------
    src_dir:
        Path to a local folder containing the report file(s).

    dst_dir:
        Path to the root network folder in which all user reports are stored.

    dst_subdir:
        Name of the network folder subdirectory to which the report file will be moved.

    Returns:
    --------
    True if report files are moved, False if no file is moved due to an error.
    """

    _logger.info("Uploading reports ...")

    for rep_path in glob(join(src_dir, "*.xlsx")):

        rep_name = split(rep_path)[1]

        # compile paths
        dst_path = join(dst_dir, dst_subdir)
        dst_file_path = join(dst_path, rep_name)
        _logger.info(f" Moving file: {rep_path} -> {dst_file_path} ...")

        # check if the destination folder exists,
        # create new subdir in the folder if not
        if not exists(dst_path):
            try:
                mkdir(dst_path)
            except FileExistsError:
                _logger.error("Could not upload reports. "
                f"Reason: Missing access to destination folder '{dst_dir}'.")
                return False
            except Exception as exc:
                _logger.error("Could not create report directory "
                             f"in destination folder. Reason: {exc}.")
                return False

        # upload report to the dst folder
        try:
            move(rep_path, dst_file_path)
        except Exception as exc:
            _logger.error(f"Moving failed. Reason: {exc}")
            return False

    return True

def _notify_users(notif_cfg: dict, countries: dict) -> bool:
    """
    Manages seding of email notifications to users.

    Params:
    -------
    notif_cfg:
        Application 'notifications' configuration params.

    countries:
        Countries and their company codes for which the \n
        reports will be created.

    Returns:
    --------
    True if sending of the notification succeeds, False if it fails.
    """

    _logger.info("Sending notification to users ...")

    subject = notif_cfg["subject"].replace("$date$" , get_current_date("%d-%b-%Y"))
    recips = []

    for usr in notif_cfg["recipients"]:
        if usr["mail"] not in recips and (usr["country"] in countries or usr["country"] == "All"):
            recips.append(usr["mail"])

    notif_path = join(notif_cfg["notification_dir"], notif_cfg["notification_name"])

    with open(notif_path, 'r', encoding = "utf-8") as stream:
        body = stream.read()

    msg = mail.create_message(notif_cfg["sender"], recips, subject, body)

    try:
        mail.send_smtp_message(msg, notif_cfg["host"], notif_cfg["port"])
    except mail.UndeliveredWarning as wng:
        _logger.warning(wng)
    except TimeoutError as exc:
        _logger.exception(exc)
        return False
    except Exception as exc:
        _logger.exception(exc)
        return False

    return True

def report_output(countries: dict, data: DataFrame,
                  report_cfg: dict, notif_cfg: dict) -> bool:
    """
    Manages reporting of the processign output to the end users.

    Params:
    -------
    countries:
        Countries and their company codes \n
        for which the reports will be created.

    data:
        Data that will be written to the Excel report.

    report_cfg:
        Application 'reports' configuration params.

    notif_cfg:
        Application 'notifications' configuration params.

    Returns:
    --------
    True if reporting succeeds, False if it fails.
    """

    if not _create_reports(data, report_cfg, notif_cfg, countries):
        return False

    uploaded = _upload_reports(
        src_dir = report_cfg["local_report_dir"],
        dst_dir = report_cfg["net_report_dir"],
        dst_subdir = get_current_date(report_cfg["net_report_subdir_format"])
    )

    if not uploaded:
        return False

    if not notif_cfg["send"]:
        _logger.warning("Sending notification to users turned off in 'appconfig.yaml'.")
    else:

        if not _create_notification(notif_cfg, report_cfg):
            return False

        if not _notify_users(notif_cfg, countries):
            return False

    return True

def remove_temp_files(dir_path: str):
    """
    Deletes all files contained
    in the application temp folder.

    Params:
    -------
    dir_path:
        Path to the folder where
        temporary files are stored.

    Returns:
    --------
    None.
    """

    file_paths = glob(join(dir_path, "**", "*.*"), recursive = True)

    if len(file_paths) == 0:
        _logger.warning("No temporary files found!")
        return

    _logger.info("Deleting temporaty data ...")

    for file_path in file_paths:
        try:
            remove(file_path)
        except Exception as exc:
            _logger.exception(exc)
