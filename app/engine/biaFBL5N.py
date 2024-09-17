# pylint: disable = C0103, W0603

"""
The 'biaFBL5N.py' module automates the standard SAP GUI FBL5N transaction in order
to load and export accounting data located on customer accounts into a plain text
file.

Version history:
1.0.20210908 - initial version
1.0.20220112 - removed dymamic layout creation upon data load.
               Data layouts will now be applied by entering a layout name
               in the transaction main search mask.
1.0.20220504 - removed unused virtual key mapping from _vkeys{}
"""

from datetime import datetime, date
from os.path import exists, isfile, split
from pyperclip import copy as copy_to_clipboard
from win32com.client import CDispatch

# custom warnings
class NoDataFoundWarning(Warning):
    """
    Warns that there are no open
    items available on account.
    """

# custom exceptions
class AbapRuntimeError(Exception):
    """
    Raised when SAP 'ABAP Runtime Error'
    occurs during communication with
    the transaction.
    """

class ConnectionLostError(Exception):
    """
    Raised when a connection to SAP
    is lost as a result of a network error.
    """

class DataWritingError(Exception):
    """
    Raised when writing of accounting
    data to file fails.
    """

class DocumentFilterError(Exception):
    """
    Raised when atempting to use
    the document filter displays
    an error message.
    """

class FolderNotFoundError(Exception):
    """
    Raised when the folder to which
    data should be exported doesn't exist.
    """

class ItemsLoadingError(Exception):
    """
    Raised when loading of open
    items fails.
    """

class SapRuntimeError(Exception):
    """
    Raised when an unhanded general SAP
    error occurs during communication with
    the transaction.
    """

class TransactionNotStartedError(Exception):
    """
    Raised when attempting to use a procedure
    before starting the transaction.
    """

_sess = None
_main_wnd = None
_stat_bar = None

# keyboard to SAP virtual keys mapping
_vkeys = {
    "Enter":        0,
    "F3":           3,
    "F8":           8,
    "F9":           9,
    "CtrlS":        11,
    "F12":          12,
    "ShiftF4":      16,
    "ShiftF12":     24,
    "CtrlF1":       25
}

def _is_error_message(sbar: CDispatch) -> bool:
    """
    Checks if a status bar message
    is an error message.
    """

    if sbar.messageType == "E":
        return True

    return False

def _is_sap_runtime_error(main_wnd: CDispatch) -> bool:
    """
    Checks if a SAP ABAP runtime error exists.
    """

    if main_wnd.text == "ABAP Runtime Error":
        return True

    return False

def _is_popup_dialog() -> bool:
    """
    Checks if the active window
    is a popup dialog window.
    """

    if _sess.ActiveWindow.type == "GuiModalWindow":
        return True

    return False

def _is_alert_message(sbar: CDispatch) -> bool:
    """
    Checks if a status bar message
    is an error message.
    """

    if sbar.messageType in ("W", "E"):
        return True

    return False

def _close_popup_dialog(confirm: bool):
    """
    Confirms or delines a pop-up dialog.
    """

    if _sess.ActiveWindow.text == "Information":
        if confirm:
            _main_wnd.SendVKey(_vkeys["Enter"]) # confirm
        else:
            _main_wnd.SendVKey(_vkeys["F12"])   # decline
        return

    btn_caption = "Yes" if confirm else "No"

    for child in _sess.ActiveWindow.Children:
        for grandchild in child.Children:
            if grandchild.Type != "GuiButton":
                continue
            if btn_caption != grandchild.text.strip():
                continue
            grandchild.Press()
            return

def _toggle_worklist(activate: bool):
    """
    Activates or deactivates the 'Use worklist' option
    in the transaction main search mask.
    """

    used = _main_wnd.FindAllByName("PA_WLKUN", "GuiCTextField").Count > 0

    if (activate and not used) or (not activate and used):
        _main_wnd.SendVKey(_vkeys["CtrlF1"])

def _open_selection_list(sel_type: str):
    """
    Opens selection list option on the FBLN main search mask,
    to which company codes can be inserted.
    """

    if sel_type == "company_codes":
        _main_wnd.findById("usr/btn%_DD_BUKRS_%_APP_%-VALU_PUSH").press()
    elif sel_type == "accounts":
        _main_wnd.findById("usr/btn%_DD_KUNNR_%_APP_%-VALU_PUSH").press()

def _set_customer_account(val: str):
    """
    Enters customer account value in string format to the
    corresonding field on the FBLN main search mask. If the
    'val' parameter is an empty string, the content of the
    field will be erased.
    """

    fld_tech_name = ""

    if _main_wnd.FindAllByName("SO_WLKUN-LOW", "GuiCTextField").Count > 0:
        fld_tech_name = "SO_WLKUN-LOW"
    elif _main_wnd.FindAllByName("DD_KUNNR-LOW", "GuiCTextField").Count > 0:
        fld_tech_name = "DD_KUNNR-LOW"

    _main_wnd.FindByName(fld_tech_name, "GuiCTextField").text = val

def _select_data_format(idx: int):
    """
    Selects data export format from the export options dialog
    based on the option index on the list.
    """
    option_wnd = _sess.FindById("wnd[1]")
    option_wnd.FindAllByName("SPOPLI-SELFLAG", "GuiRadioButton")(idx).Select()

def _set_layout(name: str):
    """
    Enters layout name into the 'Layout' field
    located on the main transaction window.
    """
    _main_wnd.findByName("PA_VARI", "GuiCTextField").text = name

def _set_company_codes(vals: list):
    """
    Enters company code values into
    the 'Company code' field list.
    """

    _main_wnd.SendVKey(_vkeys["ShiftF4"])       # clear any previous values
    s_company_codes = list(map(str, vals))
    copy_to_clipboard("\r\n".join(s_company_codes))   # copy company codes to clipboard
    _main_wnd.SendVKey(_vkeys["ShiftF12"])      # place values into the selection list
    _main_wnd.SendVKey(_vkeys["F8"])            # confirm entered values
    copy_to_clipboard("")               # clear the clipboard

def _set_from_clr_date(val: str):
    """
    Enters date value to the first 'Posting date' field
    located on the main transaction window.
    """
    _main_wnd.findByName("SO_BUDAT-LOW", "GuiCTextField").text = val

def _set_to_clr_date(val: str):
    """
    Enters date value to the last 'Posting date' field
    located on the main transaction window.
    """
    _main_wnd.findByName("SO_BUDAT-HIGH", "GuiCTextField").text = val

def _set_from_clr_date(day: date):
    """
    Enters date value to the first 'Cleared date' field
    located on the main transaction window.
    """
    val = day.strftime("%d.%m.%Y")
    _main_wnd.findByName("SO_AUGDT-LOW", "GuiCTextField").text = val

def _set_to_clr_date(day: date):
    """
    Enters date value to the last 'Cleared date' field
    located on the main transaction window.
    """
    val = day.strftime("%d.%m.%Y")
    _main_wnd.findByName("SO_AUGDT-HIGH", "GuiCTextField").text = val

def _choose_line_item_selection(option: str):
    """
    Selects the kind of items to load.
    """

    if option == "all_items":
        _main_wnd.findByName("X_AISEL", "GuiRadioButton").select()
    elif option == "open_items":
        _main_wnd.findByName("X_OPSEL", "GuiRadioButton").select()
    elif option == "cleared_items":
        _main_wnd.findByName("X_CLSEL", "GuiRadioButton").select()
    else:
        assert False, "Unrecognized selection option!"

def _apply_document_filter(kind: str):
    """
    Applies document type filter
    on the items to load.
    """

    if kind == "credit_memo":
        doctype = "DG"
    else:
        assert False, "Unrecgnized value used!"

    # open filer menu
    _main_wnd.SendVKey(_vkeys["ShiftF4"])

    # raise error for any unexpected
    # messages and let the caller handle it
    if _is_alert_message(_stat_bar):
        raise DocumentFilterError(_stat_bar.text)

    # set the search mask to load only data for credit notes
    _main_wnd.FindByName("%%DYN015-LOW", "GuiCTextField").text = doctype # credit notes

def _load_items():
    """
    Simulates pressing the 'Execute'
    button that triggers item loading.
    """

    try:
        _main_wnd.SendVKey(_vkeys["F8"])
    except Exception as exc:
        raise SapRuntimeError("Loading of accounting data failed!") from exc

    try: # SAP crash can be caught only after next statement following item loading
        msg = _stat_bar.Text
    except Exception as exc:
        raise ConnectionLostError("Connection to SAP lost due to an network error.") from exc

    if _is_sap_runtime_error(_main_wnd):
        raise SapRuntimeError("SAP runtime error!")

    if "items displayed" not in msg:
        raise NoDataFoundWarning(msg)

    if "The current transaction was reset" in msg:
        raise SapRuntimeError("FBL5N was unexpectedly terminated!")

    if _is_error_message(_stat_bar):
        raise ItemsLoadingError(msg)

    if _main_wnd.text == 'ABAP Runtime Error':
        raise AbapRuntimeError("Data loading failed due to an ABAP runtime error.")

def _export_to_file(file_path: str, enc: str = "4120"):
    """
    Exports loaded accounting data to a text file.
    """

    folder_path, file_name = split(file_path)

    if not exists(folder_path):
        raise FolderNotFoundError(f"Export folder not found: {folder_path}")

    if not file_path.endswith(".txt"):
        raise ValueError(f"Invalid file type: {file_path}. "
        "Only '.txt' file types are supported.")

    _main_wnd.SendVKey(_vkeys["F9"])     # open local data file export dialog
    _select_data_format(0)               # set plain text data export format
    _main_wnd.SendVKey(_vkeys["Enter"])  # confirm

    _sess.FindById("wnd[1]").FindByName("DY_PATH", "GuiCTextField").text = folder_path
    _sess.FindById("wnd[1]").FindByName("DY_FILENAME", "GuiCTextField").text = file_name
    _sess.FindById("wnd[1]").FindByName("DY_FILE_ENCODING", "GuiCTextField").text = enc

    _main_wnd.SendVKey(_vkeys["CtrlS"])  # replace an exiting file
    _main_wnd.SendVKey(_vkeys["F3"])     # Load main mask

    # double check if data export succeeded
    if not isfile(file_path):
        raise DataWritingError(f"Failed to export data to file: {file_path}")

def start(sess: CDispatch):
    """
    Starts FBL5N transaction.

    Params:
    ------
    sess:
        A GuiSession object.

    Returns:
    -------
    None.
    """

    global _sess
    global _main_wnd
    global _stat_bar

    _sess = sess
    _main_wnd = _sess.findById("wnd[0]")
    _stat_bar = _main_wnd.findById("sbar")

    _sess.StartTransaction("FBL5N")

def close():
    """
    Closes a running FBL5N transaction.

    Params:
    -------
    None.

    Returns:
    --------
    None.

    Raises:
    -------
    TransactionNotStartedError:
        When attempting to close
        FBL5N when it's not running.
    """

    global _sess
    global _main_wnd
    global _stat_bar

    if _sess is None:
        raise TransactionNotStartedError("Cannot close FBL5N when it's actually not running!"
        "Use the biaFBL5N.start() procedure to run the transaction first of all.")

    _sess.EndTransaction()

    if _is_popup_dialog():
        _close_popup_dialog(confirm = True)

    _sess = None
    _main_wnd = None
    _stat_bar = None

def export(file_path: str, layout: str, company_codes: list,
           from_clr_date: date = None, to_clr_date: date = None):
    """
    Exports credit notes data from customer accounts into a plain text file. \n
    If from_clr_date only is provided then all items posted from that date
    up to current date (including) will be exported. \n

    If to_clr_date only is provided then all items on accounts posted up to
    that date (including) will be exported. \n

    If to_clr_date and from_clr_date are both provided then all items on accounts
    posted from to_clr_date to from_clr_date (including) will be exported. \n

    If both to_clr_date and from_clr_date are both equal then all items on accounts
    posted on that date will be exported. \n

    If both to_clr_date and from_clr_date are None, then open items only \n
    for the current date will be exported.

    Params:
    -------
    file_path:
        Path to the file to which the data will be exported.

    layout:
        Name of the layout (default "") defining format of the loaded/exported data.

    company_codes:
        Company codes for which the data will be exported.

    from_clr_date:
        Date (including) from which data for all cleared credit notes will be exported.

    to_clr_date:
        Date (including) to which data for all cleared credit notes will be exported.

    Returns:
    --------
    None.
    """

    if _sess is None:
        raise TransactionNotStartedError("Cannot export accounting data from FBL5N "
        "when it's actually not running! Use the biaFBL5N.start() procedure to run "
        "the transaction first of all.")

    if isinstance(from_clr_date, date) and isinstance(to_clr_date, date):
        if from_clr_date > to_clr_date:
            raise ValueError("Export 'from' date cannot be greated than export 'to' date!")

    _toggle_worklist(activate = False)
    _open_selection_list("company_codes")
    _set_company_codes(company_codes)
    _set_layout(layout)
    _set_customer_account("")

    if from_clr_date is not None and to_clr_date is None:
        _set_from_clr_date(from_clr_date)
        _set_to_clr_date(datetime.now.date())
        _choose_line_item_selection("cleared_items")
    elif from_clr_date is None and to_clr_date is not None:
        _set_from_clr_date("")
        _set_to_clr_date(to_clr_date)
        _choose_line_item_selection("cleared_items")
    elif from_clr_date is not None and to_clr_date is not None:
        _set_from_clr_date(from_clr_date)
        _set_to_clr_date(to_clr_date)
        _choose_line_item_selection("cleared_items")
    else:
        _choose_line_item_selection("open_items")

    _apply_document_filter("credit_memo")
    _load_items()
    _export_to_file(file_path)
