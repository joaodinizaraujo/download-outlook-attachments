import os
import sys
from time import sleep

from src.utils.outlook_utils import (
    is_outlook_open,
    open_outlook,
    check_email
)
from src.utils.path_utils import (
    create_directory_if_not_exists,
    replace_without_overwriting
)

# diretório base, utiliza o getcwd caso seja o .exe
# no contrário utiliza a pasta da main.py
if getattr(sys, "frozen", False):
    SRC_DIR = os.getcwd()
    DOCS_DIR = os.path.join(SRC_DIR, "documentos")
else:
    SRC_DIR = os.path.dirname(__file__)
    DOCS_DIR = os.path.join(os.path.join(SRC_DIR, ".."), "documentos")

TEMP_DOCS_DIR = os.path.join(DOCS_DIR, "temp")


def main():
    create_directory_if_not_exists(DOCS_DIR,
                                   TEMP_DOCS_DIR)

    while not is_outlook_open():  # abrindo outlook
        open_outlook(is_outlook_open())
        sleep(10)

