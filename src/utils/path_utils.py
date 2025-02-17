import os
import re


def sanitize_folder_name(s: str) -> str:
    return re.sub(r"[<>:\"/\\|?*']", ' - ', s).strip()


def create_directory_if_not_exists(*dirs: str) -> None:
    """
    Cria as pastas especificadas, se elas ainda não existirem.
    :param dirs: Lista de caminhos de diretórios
    """

    for directory in dirs:
        if not os.path.exists(directory):
            os.makedirs(directory)
