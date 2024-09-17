# pylint: disable = C0103, W0603

"""
The 'biaDMS.py' module uses the standard SAP GUI UDM_DISPUTE transaction
in order to automate searching, export and updating disputed case data.

Version history:
1.0.20210526 - initial version
1.0.20210908 - removed 'srch_mask' parameter from 'close()' procedure and any related logic
             - added assertions as input check to public procedures
1.0.20220427 - fixed bug in 'modify_case_parameters()' when edit mode was incorrectly
               identified as active following an error
               during processing of a previous case.
1.0.20220504 - removed unused virtual key mapping from _vkeys{}
"""

from enum import Enum, IntEnum
import logging
from os.path import exists, isfile, split
from pyperclip import copy as copy_to_clipboard
from win32com.client import CDispatch

class CaseCountError(Exception):
    """
    When the number of searched cases
    exceeds the maximum 5000 entries.
    """

class CaseEditingError(Exception):
    """
    When attempting to edit case
    parameters results in an error.
    """

class DataWritingError(Exception):
    """
    Raised when writing of accounting
    data to file fails.
    """

class TransactionNotStartedError(Exception):
    """
    Raised when attempting to use a procedure
    before starting the transaction.
    """

class LayoutNotFoundError(Exception):
    """
    Raised when the data layout
    used doesn't exist.
    """

class FolderNotFoundError(Exception):
    """
    Raised when the folder to which
    data should be exported doesn't exist.
    """

class CaseStates(IntEnum):
    """
    Available DMS case 'Status' values.
    """
    Original = 0
    Open = 1
    Solved = 2
    Closed = 3
    Devaluated = 4

class RootCauses(Enum):
    """
    Available DMS case 'Root Cause Code' values.
    """
    UNJUSTIFIED_DISPUTE = "L00"
    PAYMENT_AGREMENT ="L01"
    PAYMENT_TERMS_STRETCH = "L04"
    CREDITNOTE_ISSUED = "L06"
    CHARGE_OFF = "L08"
    CLOSED_WHILE_UNDER_50 = "L14"

class _SearchFieldIndexes(IntEnum):
    """
    Indexes of DMS search mask fields.
    """
    CaseID = 0
    RestrictHitsTo = 23


_sess = None
_main_wnd = None
_stat_bar = None

# keyboard to SAP virtual keys mapping
_vkeys = {
    "Enter":    0,
    "F3":       3,
    "F8":       8,
    "CtrlS":    11,
    "F12":      12,
    "ShiftF4":  16,
    "ShiftF12": 24
}

_status_map = {
    CaseStates.Open: "Open",
    CaseStates.Solved: "Solved",
    CaseStates.Closed: "Closed",
}

_logger = logging.getLogger("master")

def _is_error_message(sbar: CDispatch) -> bool:
    """
    Checks if a status bar message
    is an error message.
    """

    if sbar.messageType == "E":
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

def _get_grid_view() -> CDispatch:
    """
    Returns a GuiGridView object representing
    the DMS window containing search results.
    """

    splitter_shell = _main_wnd.FindByName("shell", "GuiSplitterShell")
    grid_view = splitter_shell.FindAllByName("shell", "GuiGridView")(6)

    return grid_view

def _get_param_mask() -> CDispatch:
    """
    Returns a GuiGridView object representing
    the DMS case parameter mask containing editable fields.
    """

    splitter_shell = _main_wnd.FindByName("shell", "GuiSplitterShell")
    param_mask = splitter_shell.FindAllByName("shell", "GuiGridView")(5)

    return param_mask

def _execute_query():
    """
    Simulates pressing the 'Search' button
    located on the DMS main search mask.
    """

    splitter_shell = _main_wnd.FindByName("shell", "GuiSplitterShell")
    qry_toolbar = splitter_shell.FindAllByName("shell", "GuiToolbarControl")(5)
    qry_toolbar.PressButton("DO_QUERY")

def _find_and_click_node(tree: CDispatch, node: CDispatch, node_id: str) -> bool:
    """
    Traverses the left-sided DMS menu tree to find the item with the given node ID.
    Once the item is found, the procedure simulates clicking on that item to open
    the corresponding subwindow.
    """

    # find and double click the target root node
    if tree.IsFolder(node):
        tree.CollapseNode(node)
        tree.ExpandNode(node)

    # double clisk the target node
    if node.strip() == node_id:
        tree.DoubleClickNode(node)
        return True

    subnodes = tree.GetsubnodesCol(node)

    if subnodes is None:
        return False

    iter_subnodes = iter(subnodes)

    if _find_and_click_node(tree, next(iter_subnodes), node_id):
        return True

    try:
        next_node = next(iter_subnodes)
    except StopIteration:
        return False
    else:
        return _find_and_click_node(tree, next_node, node_id)

def _get_search_mask() -> CDispatch:
    """
    Returns the GuiGridView object representing
    the DMS case search window.
    """

    # find the target node by traversing the search tree
    tree = _main_wnd.findById(
        "shellcont/shell/shellcont[0]/shell/shellcont[1]/shell/shellcont[1]/shell"
    )

    nodes = tree.GetNodesCol()
    iter_nodes = iter(nodes)
    clicked = _find_and_click_node(tree, next(iter_nodes), node_id = "4")

    assert clicked, "Target node not found!"

    # get reference to the search mask object found
    splitter_shell = _main_wnd.FindByName("shell", "GuiSplitterShell")
    srch_mask = splitter_shell.FindAllByName("shell", "GuiGridView")(4)

    return srch_mask

def _change_case_param(param_mask: CDispatch, cell_idx: int, col_type: str, val: str):
    """
    Changes the value of an editable case parameter field identified
    in the case parameter grid by the cell index and column type.
    """

    param_mask.ModifyCell(cell_idx, col_type, val)

def _get_case_param(param_mask: CDispatch, cell_idx: int, col_type: str) -> str:
    """
    Returns the value of an editable case parameter field identified
    in the case parameter grid by the cell index and column type.
    """

    cell_val = param_mask.GetCellValue(cell_idx, col_type)

    return cell_val

def _get_control_toolbar() -> CDispatch:
    """
    Returns GuiToolbarControl object representing the DMS control toolbar
    located in the transaction upper window.
    """

    splitter_shell = _main_wnd.FindByName("shell", "GuiSplitterShell")
    ctrl_toolbar = splitter_shell.FindAllByName("shell", "GuiToolbarControl")(3)

    return ctrl_toolbar

def _save_changes():
    """
    Simulates pressing the 'Save' button located
    in the DMS upper toolbar.
    """

    _get_control_toolbar().PressButton("SAVE")

    err_msg = ""

    if _is_error_message(_stat_bar):
        err_msg = _stat_bar.Text
        return (False, err_msg)

    return (True, err_msg)

def _change_case_status(param_mask: CDispatch, val: int):
    """
    Changes the case 'Status' parameter.
    """

    prev_val = _get_case_param(param_mask, 0, "VALUE2")
    new_val = _status_map[val]

    while prev_val != new_val:

        if prev_val == _status_map[CaseStates.Open]:
            curr_val = _status_map[CaseStates.Solved]
            _change_case_param(param_mask, 0, "VALUE2", curr_val)

            if new_val == _status_map[CaseStates.Closed]:
                _save_changes()

        elif prev_val == _status_map[CaseStates.Closed]:
            curr_val = _status_map[CaseStates.Solved]
            _change_case_param(param_mask, 0, "VALUE2", curr_val)

            if new_val == _status_map[CaseStates.Open]:
                _save_changes()

        elif prev_val == _status_map[CaseStates.Solved]:
            curr_val = new_val
            _change_case_param(param_mask, 0, "VALUE2", curr_val)

        prev_val = curr_val

def _apply_layout(grid_view: CDispatch, name: str):
    """
    Searches a layout by name in the DMS layouts list. If the layout is
    found in the list of available layouts, this gets selected.
    """

    # Open Change Layout Dialog
    grid_view.PressToolbarContextButton("&MB_VARIANT")
    grid_view.SelectContextMenuItem("&LOAD")
    apo_grid = _sess.findById("wnd[1]").findAllByName("shell", "GuiShell")(0)

    for row_idx in range(0, apo_grid.RowCount):
        if apo_grid.GetCellValue(row_idx, "VARIANT") == name:
            apo_grid.setCurrentCell(str(row_idx), "TEXT")
            apo_grid.clickCurrentCell()
            return

    raise LayoutNotFoundError(f"Layout not found: {name}")

def _export_to_file(grid_view: CDispatch, file_path: str, enc: str = "4120"):
    """
    Enters folder path, file name and encoding of the file
    to which the exported data will be written.
    """

    if not file_path.endswith(".txt"):
        raise ValueError(
            f"Invalid file type: {file_path}. "
            "Only '.txt' file types are supported."
        )

    folder_path, file_name = split(file_path)

    if not exists(folder_path):
        raise FolderNotFoundError(f"Export folder not found: {folder_path}")

    grid_view.PressToolbarContextButton("&MB_EXPORT")
    grid_view.SelectContextMenuItem("&PC")
    _sess.FindById("wnd[1]").FindAllByName("SPOPLI-SELFLAG", "GuiRadioButton")(0).Select()
    _main_wnd.SendVKey(_vkeys["Enter"])

    _sess.FindById("wnd[1]").FindByName("DY_PATH", "GuiCTextField").text = folder_path
    _sess.FindById("wnd[1]").FindByName("DY_FILENAME", "GuiCTextField").text = file_name
    _sess.FindById("wnd[1]").FindByName("DY_FILE_ENCODING", "GuiCTextField").text = enc

    _main_wnd.SendVKey(_vkeys["CtrlS"])  # replace an exiting file
    _main_wnd.SendVKey(_vkeys["F3"])     # Load main mask

    # double check if data export succeeded
    if not isfile(file_path):
        raise DataWritingError(f"Failed to export data to file: {file_path}")

def _toggle_display_change(activate: bool) -> tuple:
    """
    Enables editing of a case.
    """

    _get_control_toolbar().PressButton("TOGGLE_DISPLAY_CHANGE")

    msg = _stat_bar.Text

    if "display only" in msg:
        return (False, msg)

    if not activate:
        _main_wnd.SendVKey(_vkeys["F3"])

    # handle alert dialogs for non-editable cases
    if _is_popup_dialog():

        err_msg = _sess.ActiveWindow.children(1).children(1).text

        if err_msg == "Attributes may be overwritten later":
            _close_popup_dialog(confirm = True)
        else:
            _close_popup_dialog(confirm = False)
            _main_wnd.SendVKey(_vkeys["F3"])
            return (False, err_msg)

    elif _is_error_message(_stat_bar):
        err_msg = _stat_bar.Text
        return (False, err_msg)

    return (True, "")

def start(sess: CDispatch) -> CDispatch:
    """
    Starts UDM_DISPUTE transaction.

    Params:
    ------
    sess:
        A GuiSession object.

    Returns:
    -------
    A GuiGridView object representing the search window.
    """

    global _sess
    global _main_wnd
    global _stat_bar

    _sess = sess
    _main_wnd = _sess.findById("wnd[0]")
    _stat_bar = _main_wnd.findById("sbar")

    _sess.StartTransaction("UDM_DISPUTE")
    srch_mask = _get_search_mask()

    return srch_mask

def close():
    """
    Closes a running UDM_DISPUTE transaction.

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
        UDM_DISPUTE when it's not running.
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

def _get_nfound(msg: str) -> int:

    num = msg.split(" ")[0]
    num = num.strip().replace(".", "")

    return int(num)

def _set_case(search_mask, case: int):
    search_mask.ModifyCell(_SearchFieldIndexes.CaseID, "VALUE1", str(case))

def _set_cases(search_mask: CDispatch, cases: list):

    search_mask.PressButton(0, "SEL_ICON1")
    _main_wnd.SendVKey(_vkeys["ShiftF4"])               # clear any previous values
    copy_to_clipboard("\r\n".join(map(str, cases)))     # copy accounts to clipboard
    _main_wnd.SendVKey(_vkeys["ShiftF12"])              # confirm selection
    copy_to_clipboard("")                               # clear the clipboard
    _main_wnd.SendVKey(_vkeys["F8"])                    # confirm

def _set_hitlimit(search_mask, n_cases):
    search_mask.ModifyCell(_SearchFieldIndexes.RestrictHitsTo, "VALUE1", n_cases)

def search_dispute(search_mask: CDispatch, case: int) -> CDispatch:
    """
    Searches a disputed case in DMS based on the case ID.

    Params:
    -------
    search_mask:
        An instantiated GuiGridView object representing DMS search mask window.

    case:
        ID number of the disputed case.

    Returns:
    --------
    A GuiGridView object representing the search result
    """

    _set_case(search_mask, case)
    _execute_query()

    n_found = _get_nfound(_stat_bar.Text)
    item_list = None

    if n_found > 0:
        item_list = _get_grid_view()

    return item_list

def search_disputes(search_mask: CDispatch, cases: tuple) -> tuple:
    """
    Searches disputed cases based in DMS on their database IDs.

    Params:
    -------
    search_mask:
        A GuiGridView object representing the DMS case search window.

    cases:
        ID numbers (strings or integers) of the disputed case.

    Returns:
    --------
    A tuple of an instance of GuiGridView object representing
    the search result list grid and the number of cases found.

    Raises:
    -------
    CaseCountError:
        When the number of searched cases is mre than 5000.
    """

    MAX_DISPUTES = 5000

    for case in cases:
        if not str(case).isnumeric():
            raise ValueError(f"Incorrect case value found: {case}")

    if len(cases) > MAX_DISPUTES:
        raise CaseCountError(f"The maximum limit of 5000 cases exceeded: {len(cases)}")

    _set_hitlimit(search_mask, MAX_DISPUTES)
    _set_cases(search_mask, cases)
    _execute_query()

    n_found = _get_nfound(_stat_bar.Text)
    item_list = None

    if n_found > 0:
        item_list = _get_grid_view()

    return (item_list, n_found)

def modify_case_parameters(grid_view: CDispatch, root_cause: str = None, status_sales: str = None,
                           status: CaseStates = CaseStates.Original):
    """
    Modifies parameters of a disputed case.

    Params:
    -------
    grid_view:
        An instantiated GuiGridView object representing \n
        DMS window containing case search results.

    root_cause:
        Represents 'Root Cause Code' parameter of a disputed case.

    status_sales:
        Represents 'Status Sales' parameter of a disputed case.

    stat:
        Represents 'Status' parameter of a disputed case.

    Returns:
    --------
    None.

    Raises:
    -------
    CaseEditingError:
        When attempting to change case parameters fails.
    """

    if not (root_cause is None or isinstance(root_cause, RootCauses)):
        raise TypeError(f"Argument 'root_cause' has incorrect type: {type(root_cause)}")

    if not (status is None or isinstance(status, CaseStates)):
        raise TypeError(f"Argument 'status' has incorrect type: {type(status)}")

    MAX_FIELD_CHARS = 40

    if status_sales is not None and len(status_sales) > MAX_FIELD_CHARS:
        ValueError("The limit of 50 characters for 'Status Sales' exceeded!")

    # open case details
    grid_view.DoubleClickCurrentCell()

    # enter edit mode
    activated, err_msg = _toggle_display_change(activate = True)

    if not activated:
        _main_wnd.SendVKey(_vkeys["F3"])
        raise CaseEditingError(err_msg)

    param_mask = _get_param_mask()

    if root_cause is not None:
        _change_case_param(param_mask, 10, "VALUE2", root_cause.value)

    if status != CaseStates.Original:
        _change_case_status(param_mask, status.value)

    if status_sales is not None:
        _change_case_param(param_mask, 11, "VALUE1", status_sales)

    saved, saving_msg = _save_changes()

    if not saved:
        _main_wnd.SendVKey(_vkeys["F3"])
        if _is_popup_dialog():
            _close_popup_dialog(confirm = False)
        raise CaseEditingError(saving_msg)

    # exit edit mode
    deactivated, displaying_msg = _toggle_display_change(activate = False)

    if not deactivated:
        raise CaseEditingError(displaying_msg)

def export(grid_view: CDispatch, file_path: str, layout: str):
    """
    Exports disputed data into a plain text file.

    Params:
    -------
    grid_view:
        An instantiated GuiGridView object representing \n
        the DMS window containing case search results.

    file_path:
        Path to the text file to which data will be exported.

    layout:
        Name of the layout conaining data fields to export.

    Returns:
    --------
    None.
    """

    _apply_layout(grid_view, layout)
    _export_to_file(grid_view, file_path)
