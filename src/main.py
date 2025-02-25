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
    GPT_KEY_FILE = os.path.join(SRC_DIR, "key.txt")
else:
    SRC_DIR = os.path.dirname(__file__)
    ROOT_DIR = os.path.join(SRC_DIR, "..")
    DOCS_DIR = os.path.join(ROOT_DIR, "docs")
    GPT_KEY_FILE = os.path.join(ROOT_DIR, "key.txt")


def main():
    print("Começando...")

    openai_key = None
    if not os.path.exists(GPT_KEY_FILE):
        print("Chave do GPT não encontrada... Não terá resumos rs")
    else:
        openai_key = open(GPT_KEY_FILE).read()

    create_directory_if_not_exists(DOCS_DIR)

    while not is_outlook_open():  # abrindo outlook
        open_outlook(is_outlook_open())
        sleep(10)

    data = check_email(DOCS_DIR, openai_key)  # pegando dados dos emails

    if len(data) > 0:
        print("Anexos baixados!")
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
