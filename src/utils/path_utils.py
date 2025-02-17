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


def join_without_overwriting(base_dir: str, file_name: str) -> str:
    """
    Gera um caminho único para o arquivo, evitando sobrescrita.
    :param base_dir: diretório onde o arquivo será salvo
    :param file_name: nome do arquivo a ser salvo
    :return: caminho do arquivo com nome único
    """

    base_name, extension = os.path.splitext(file_name)
    counter = 1
    path = os.path.join(base_dir, file_name)

    while os.path.exists(path):
        new_file = f"{base_name}_{counter}{extension}"
        path = os.path.join(base_dir, new_file)
        counter += 1

    return path
