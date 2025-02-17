import os
from typing import Any

import psutil
import pyautogui as pg
import win32com.client as client

from src.utils.path_utils import join_without_overwriting


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

    for i in range(1, attachments.Count + 1):
        attachment = attachments.Item(i)
        cont += 1

        saved_file_path = join_without_overwriting(xml_dir, attachment.FileName)

        try:
            attachment.SaveAsFile(saved_file_path)
        except Exception as e:
            print(f"Erro ao salvar '{attachment.FileName}': {e}")

    if cont == 0:
        return False

    return True


# def get_folder_by_name(folder_name: str):
#     """
#     Retorna uma pasta no nível raiz do Outlook pelo nome.
#     :param folder_name: Nome da pasta desejada
#     :return: Objeto Folder do Outlook ou None se não encontrada
#     """
#     outlook = client.Dispatch("Outlook.Application")
#     namespace = outlook.GetNamespace("MAPI")

#     for account in namespace.Folders:
#         if folder_name in [folder.Name for folder in account.Folders]:
#             return account.Folders[folder_name]

#     return None



def get_inbox():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    return inbox

def check_email(docs_dir: str) -> list[dict[str, str]]:
    """
    Varre emails não lidos na Caixa de Entrada e salva os documentos.
    :param docs_dir: pasta onde os documentos serão salvos
    :return: lista de dicionários contendo as principais informações das mensagens processadas
    """
    inbox = get_inbox()
    data = []
    
    for message in inbox.Items:
        try:
            save_attachments(message, docs_dir)
        except Exception as e:
            print(f"Erro ao processar anexos: {e}")

        message_data = get_email_info(message)
        data.append(message_data)
    
    return data
