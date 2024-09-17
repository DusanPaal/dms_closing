"""Debugging interface"""

import logging
import os

import cv2
import numpy as np
import pyautogui

_logger = logging.getLogger("master")

def take_screenshot(img_folder: str = None, max_screens: int = None):
    """Takes a screenshot of the windows desktop."""

    if img_folder is None:
        img_folder = os.path.split(__file__)[0]

    n_img = 1
    max_screens = 999 if max_screens is None else max_screens
    n_places = len(str(max_screens))

    idx = str(n_img).zfill(n_places)
    img_name = f"screen_{idx}.png"
    img_path = os.path.join(img_folder, img_name)

    while os.path.isfile(img_path):
        n_img += 1
        idx = str(n_img).zfill(n_places)
        img_name = f"screen_{idx}.png"
        img_path = os.path.join(img_path, img_name)

    _logger.info("Taking screenhot...")
    image = pyautogui.screenshot()
    _logger.info("Creating image...")
    image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
    _logger.info("Writing image...")
    cv2.imwrite(img_path, image)
    _logger.info("Image written.")
