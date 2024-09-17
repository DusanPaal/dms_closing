# pylint: disable = C0103, E0611

"""
The 'biaSAP.py' module uses win32com (package: pywin32) to connect
to SAP GUI scripting engine and login/logout to/from the defined
SAP system.
"""

from enum import Enum
from os.path import isfile
from subprocess import Popen, TimeoutExpired
import win32com.client
from win32ui import FindWindow
from win32ui import error as WinError
from win32com.client import CDispatch

class Systems(Enum):
    """
    SAP systems and their full names.
    """
    P25 = "OG ERP: P25 Productive SSO"
    Q25 = "OG ERP: Q25 Quality Assurance SSO"

class LoginError(Exception):
    """
    Raised when logign to the SAP
    GUI scriptng engine fails.
    """

def _window_exists(name: str) -> bool:
    """Checks wheter SAP GUI process is running."""

    try:
        FindWindow(None, name)
    except WinError:
        return False
    else:
        return True

def _start_process(exe_path: str):
    """Starts a new SAP GUI process."""

    TIMEOUT = 8

    try:
        proc = Popen(exe_path)
        proc.communicate(timeout = TIMEOUT)
    except TimeoutExpired:
        pass
    except Exception as exc:
        raise LoginError("Communication with the process failed!") from exc

def login(sap_path: str, system: Systems) -> None:
    """
    Logs into the SAP GUI application.

    Params:
    -------
    sap_path:
        Path to the SAP GUI executable.

    system:
        SAP system to which connection will be created.

    Returns:
    --------
    An initialized SAP GuiSession object.

    Rasies:
    -------
    LoginError:
        Raised when logign to the SAP
        GUI scriptng engine fails.
    """

    if not isfile(sap_path):
        raise FileNotFoundError("SAP GUI executable "
        f"not found at the specified path: {sap_path}!")

    if not _window_exists("SAP Logon 750"):
        _start_process(sap_path)

    try:
        sap_gui_auto = win32com.client.GetObject('SAPGUI')
    except Exception as exc:
        raise LoginError("Could not get the 'SAPGUI' object.") from exc

    engine = sap_gui_auto.GetScriptingEngine

    if engine.Connections.Count == 0:
        engine.OpenConnection(system.value, Sync = True)

    conn = engine.Connections(0)
    sess = conn.Sessions(0)

    return sess

def logout(sess: CDispatch):
    """
    Disconnects from SAP GUI system.

    Params:
    -------
    sess:
        A SAP GUiSession object.

    Returns:
    --------
    None.
    """

    if sess is None:
        raise ValueError("Trying to close a connection that is actually not open!")

    conn = sess.Parent
    conn.CloseSession(sess.ID)
    conn.CloseConnection()


# Private Sub HandleRepeatedConnection(ByRef Session As GuiSession)

#     If Not Session.FindById("wnd[1]/usr/radMULTI_LOGON_OPT1", False) Is Nothing Then
#         Call Session.FindById("wnd[1]/usr/radMULTI_LOGON_OPT1").Select
#         Call Session.FindById("wnd[1]/tbar[0]/btn[0]").Press
#         Call Application.Wait(Now + TimeValue("0:00:01"))
#     End If

# End Sub