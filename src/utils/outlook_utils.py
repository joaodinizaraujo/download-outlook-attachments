import os
from typing import Any

import psutil
import pyautogui as pg
import win32com
import win32com.client as client
from datetime import datetime

from src.utils.path_utils import create_directory_if_not_exists, sanitize_folder_name

NOT_ACCEPTED_FORMATS = [
    ".png",
    ".jpg",
    ".jpeg",
    ".ics"
]
MAX_PATH_LENGTH = 255


def get_email_info(message: Any) -> dict[str, str]:
    """
    :param message: mensagem do outlook
    :return: dict com principais informações da mensagem
    """

    return {
        "subject": message.Subject,
        "body": message.Body,
        "sender_email": message.SenderEmailAddress,
        "sender_name": message.SenderName,
        "attachments": message.attachments
    }


def is_outlook_open() -> bool:
    """
    :return: bool indicando existência do processo do outlook
    """

    for i in psutil.process_iter():
        try:
            if "outlook" in i.name().lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

    return False


def open_outlook(is_process_open: bool) -> None:
    """
    :param is_process_open: bool indicando se o processo do outlook está aberto
    """

    if is_process_open:
        for window in pg.getAllWindows():
            titulo = window.title

            if "Segurança do Windows" in titulo:
                password = ""
                while len(password) == 0:
                    password = input("Digite sua senha do email: ").strip()
                    print("Senha digitada incorreta. Tente novamente.")

                window.activate()
                pg.typewrite(password, interval=0.1)
                pg.press("enter")
    else:
        print("Abrindo outlook...")
        os.startfile("outlook")


def save_attachments(message: Any, docs_dir: str) -> bool:
    """
    Varre os anexos da mensagem e salva os documentos sem sobrescrita.
    :param message: mensagem do Outlook
    :param docs_dir: pasta onde os documentos serão salvos
    :return: bool indicando se salvou algum anexo
    """
    cont = 0
    attachments = message.Attachments
    info = get_email_info(message)

    for i in range(1, attachments.Count + 1):
        attachment = attachments.Item(i)
        cont += 1

        filename, extension = os.path.splitext(attachment.FileName)
        if extension.lower() in NOT_ACCEPTED_FORMATS or len(extension) == 0:
            continue

        sender_folder = sanitize_folder_name(info["sender_name"])
        subject_folder = sanitize_folder_name(info["subject"])
        base_path = os.path.join(docs_dir, sender_folder, subject_folder)

        if len(base_path) > MAX_PATH_LENGTH - 50:
            subject_folder = subject_folder[:MAX_PATH_LENGTH - len(docs_dir) - len(sender_folder) - 10]
            base_path = os.path.join(docs_dir, sender_folder, subject_folder)

        create_directory_if_not_exists(base_path)

        saved_file_path = os.path.join(base_path, attachment.FileName)

        if len(saved_file_path) > MAX_PATH_LENGTH:
            max_filename_length = MAX_PATH_LENGTH - len(base_path) - len(extension) - 5
            filename = filename[:max_filename_length] + "_cut"
            saved_file_path = os.path.join(base_path, filename + extension)

        try:
            attachment.SaveAsFile(saved_file_path)
        except Exception as e:
            print(f"Erro ao salvar '{attachment.FileName}': {e}")

    return cont > 0


def get_inbox():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    return inbox


def check_email(base_dir: str) -> list[dict[str, str]]:
    """
    Varre emails não lidos na Caixa de Entrada e salva os documentos.
    :param base_dir: pasta onde os documentos serão salvos
    :return: lista de dicionários contendo as principais informações das mensagens processadas
    """
    inbox = get_inbox()
    data = []
    current_year = datetime.now().year

    for message in inbox.Items:
        try:
            email_year = message.ReceivedTime.year
            if email_year != current_year:
                continue

            save_attachments(message, base_dir)
            message_data = get_email_info(message)
            data.append(message_data)
        except Exception as e:
            print(f"Erro ao processar e-mail: {e}")

    return data
