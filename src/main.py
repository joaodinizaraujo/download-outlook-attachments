import os
import sys
from time import sleep

from src.utils.outlook_utils import (
    is_outlook_open,
    open_outlook,
    check_email
)
from src.utils.path_utils import (
    create_directory_if_not_exists
)

# diretório base, utiliza o getcwd caso seja o .exe
# no contrário utiliza a pasta da main.py
if getattr(sys, "frozen", False):
    SRC_DIR = os.getcwd()
    DOCS_DIR = os.path.join(SRC_DIR, "docs")
else:
    SRC_DIR = os.path.dirname(__file__)
    DOCS_DIR = os.path.join(os.path.join(SRC_DIR, ".."), "docs")


def main():
    create_directory_if_not_exists(DOCS_DIR)

    while not is_outlook_open():  # abrindo outlook
        open_outlook(is_outlook_open())
        sleep(10)

    data = check_email(DOCS_DIR)  # pegando dados dos emails

    if len(data) > 0:
        print("\nAnexos baixados!")
    else:
        print("Nenhum email.")


if __name__ == "__main__":
    try:
        main()
    except (FileNotFoundError, ValueError):
        ...
    except Exception as e:
        print(f"Erro não esperado: {e}")

    input("\nPressione Enter para sair...")
