import os


def create_directory_if_not_exists(*dirs: str) -> None:
    """
    Cria as pastas especificadas, se elas ainda não existirem.
    :param dirs: Lista de caminhos de diretórios
    """

    for directory in dirs:
        if not os.path.exists(directory):
            os.makedirs(directory)


def replace_without_overwriting(origin_path: str, destiny_path: str) -> None:
    """
    Move um arquivo para o destino sem sobrescrever arquivos existentes.
    Caso o destino já tenha um arquivo com o mesmo nome, adiciona um sufixo numérico ao nome
    antes da extensão até encontrar um nome disponível.

    :param origin_path: caminho completo do arquivo de origem
    :param destiny_path: caminho completo do arquivo de destino
    """

    base_name, extension = os.path.splitext(destiny_path)
    counter = 1

    while os.path.exists(destiny_path):
        destiny_path = f"{base_name}_{counter}{extension}"
        counter += 1

    os.replace(origin_path, destiny_path)


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
