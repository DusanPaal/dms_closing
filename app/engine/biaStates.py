# pylint: disable = C0103, C0301, W0703, W1203

"""
The 'biaStates.py' module provides
user interface allowing to modify
parameters defined in states.json.
"""

from datetime import datetime as dt
from os.path import join
import sys
from biaController import load_app_config, save_states

def get_user_input() -> str:
    """
    Returns a string date
    used to modify app
    'last_run' state.

    Params:
    -------
    None.

    Returns:
    --------
    None.
    """

    stat_type = input("Select state to modify (d = last run date):")

    while stat_type.lower() == 'd':

        val = input("Enter date in format 'yyyy-mm-dd' (e.g. 2022-01-24):")

        if val.lower() == 'q':
            return None

        try:
            dt.strptime(val, "%Y-%m-%d")
        except Exception:
            print("Invalid value entered!")
            continue

        return val

    # any other state type entered by user will be ignored for now
    print("Invalid state type entered!")

    return None

def set_state() -> int:
    """
    Prompts user for new state value,
    then saves this value the respective file.

    Params:
    -------
    None.

    Returns:
    --------
    None.
    """

    cfg = load_app_config(
        cfg_path = join(sys.path[0], "appconfig.yaml"),
        states_path = join(sys.path[0], "states.json")
    )

    if cfg is None:
        return

    val = get_user_input()

    if val is None:
        print("Quitting application ...")
        return

    save_states(cfg["states"], {"last_run": val})

    print(f"State set to '{val}'.")

if __name__ == "__main__":
    set_state()
