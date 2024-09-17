# pylint: disable = C0103, R0911, R0912, R0915, W1203

"""The 'app.py' module represents the main script
of the application that contains program entry
procedure called at the script import.

Version history:
----------------
1.0.20210728 - initial version
1.0.20220114 - refactored version
1.1.20220513 - upgraded with closing of DMS cases for already cleared credit notes
             - minor bugfixes
             - modified case evaluation criteria
             - updated docstrings
"""

import logging
from os.path import join
import sys
import engine.biaController as ctrlr

log = logging.getLogger("master")

def main() -> int:
    """Program entry point.

    Returns:
    --------
    Program completion state.
    """

    logger_ok = ctrlr.initialize_log(
        cfg_path = join(sys.path[0], "logging.yaml"),
        log_path = join(sys.path[0], "log.log"),
        debug = False,
        header = {
            "Application name": "CS DMS Closing",
            "Application version": "1.1.20220513",
            "Log date": ctrlr.get_current_date("%d-%b-%Y")
        }
    )

    if not logger_ok:
        return 1

    log.info("=== Initialization ===")
    cfg = ctrlr.load_app_config(
        cfg_path = join(sys.path[0], "appconfig.yaml"),
        states_path = join(sys.path[0], "states.json")
    )

    if cfg is None:
        return 2

    rules = ctrlr.load_closing_rules(join(sys.path[0], "rules.yaml"))

    if rules is None:
        return 3

    countries = ctrlr.get_active_countries(rules)

    if countries is None:
        return 4

    sess = ctrlr.connect_to_sap(cfg["sap"])

    if sess is None:
        return 5

    log.info("=== Success ===\n")

    log.info("=== Data processing ===")
    if not ctrlr.export_fbl5n_data(cfg["data"], cfg["sap"], cfg["states"], countries, sess):
        log.info("=== Failure ===\n")
        log.info("=== Cleanup ===")
        ctrlr.disconnect_from_sap(sess)
        log.info("=== Success ===\n")
        return 6

    fbl5n_data = ctrlr.preprocess_fbl5n_data(cfg["data"], rules, countries)

    if fbl5n_data is None:
        log.info("=== Failure ===\n")
        log.info("=== Cleanup ===")
        ctrlr.disconnect_from_sap(sess)
        log.info("=== Success ===\n")
        return 7

    if not ctrlr.export_dms_data(cfg["data"], cfg["sap"], fbl5n_data, sess):
        log.info("=== Failure ===\n")
        log.info("=== Cleanup ===")
        ctrlr.disconnect_from_sap(sess)
        log.info("=== Success ===\n")
        return 8

    dms_data = ctrlr.preprocess_dms_data(cfg["data"])
    closing_input, compacted = ctrlr.process_data(
        fbl5n_data, dms_data, countries.keys(), rules)
    log.info("=== Success ===\n")

    log.info("=== Case processing ===")
    if closing_input is None:
        output = compacted
        log.warning("No items to process found.")
    else:

        output = ctrlr.check_output(cfg["data"])

        if output is None:
            output = ctrlr.process_disputes(closing_input, compacted, sess)

        if output is None:
            log.info("=== Failure ===\n")
            log.info("=== Cleanup ===")
            ctrlr.disconnect_from_sap(sess)
            log.info("=== Success ===\n")
            return 9

    log.info("=== Success ===\n")

    log.info("=== Reporting ===")
    reported = ctrlr.report_output(countries,
        data = compacted,
        report_cfg = cfg["reports"],
        notif_cfg = cfg["notifications"]
    )

    if not reported:
        log.info("=== Failure ===\n")
        log.info("=== Cleanup ===")
        ctrlr.disconnect_from_sap(sess)
        log.info("=== Success ===\n")
        return 10

    log.info("=== Success ===\n")

    log.info("=== Cleanup ===")
    ctrlr.save_states(join(sys.path[0], "states.json"))
    ctrlr.disconnect_from_sap(sess)
    ctrlr.remove_temp_files(cfg["data"]["temp_dir"])
    log.info("=== Success ===\n")

    return 0

if __name__ == "__main__":
    ret_code = main()
    log.info(f"=== System shutdown with return code: {ret_code} ===")
    logging.shutdown()
    sys.exit(ret_code)
