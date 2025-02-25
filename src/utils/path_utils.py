import os
import re


def sanitize_folder_name(s: str) -> str:
    return re.sub(r"[<>:\"/\\|?*']", ' - ', s).strip()


def create_directory_if_not_exists(*args: str) -> None:
    """
    Cria as pastas especificadas, se elas ainda não existirem.
    :param dirs: Lista de caminhos de diretórios
    """

    for directory in args:
        if not os.path.exists(directory):
            os.makedirs(directory)


def join_without_overwriting(*args, file_name: str) -> str:
    """
    Gera um caminho único para o arquivo, evitando sobrescrita.
    :param file_name: nome do arquivo a ser salvo
    :return: caminho do arquivo com nome único
    """

    base_name, extension = os.path.splitext(file_name)
    counter = 1
    path = os.path.join(*args, file_name)

    while os.path.exists(path):
        new_file = f"{base_name}_{counter}{extension}"
        path = os.path.join(*args, new_file)
        counter += 1

    return path
