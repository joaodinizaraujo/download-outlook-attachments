import os
from typing import Any

import psutil
import pyautogui as pg
import win32com
import win32com.client as client
from datetime import datetime
import tkinter as tk

from src.utils.path_utils import create_directory_if_not_exists, sanitize_folder_name, join_without_overwriting
from src.utils.doc_reader import read_pdf, read_docx
from src.utils.openai_client import send_prompt

NOT_ACCEPTED_FORMATS = [
    ".png",
    ".jpg",
    ".jpeg",
    ".ics"
]
MAX_PATH_LENGTH = 255

BNCC = ["João", "Eduardo", "Leda"]
TECH = ["Nisflei", "Modolo", "Santi", "Carlos", "Myrna", "Alex"]


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


def save_attachments(message: Any, selected_option, docs_dir: str, openai_key: str | None = None) -> bool:
    """
    Varre os anexos da mensagem e salva os documentos sem sobrescrita.
    :param openai_key: chave da api da openai para resumos
    :param message: mensagem do Outlook
    :param docs_dir: pasta onde os documentos serão salvos
    :param option: opção de salvamento (1 = normal, 2 = BNCC, 3 = TECH)
    :return: bool indicando se salvou algum anexo
    """
    BNCC = ["joão", "eduardo", "leda"]
    TECH = ["nisflei", "modolo", "santi", "carlos", "myrna", "alex"]
    
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
        #Verifica qual opção foi selecionada e faz o direcionamento correto
        if selected_option == 1:
            base_path = os.path.join(docs_dir, sender_folder, subject_folder)
        elif selected_option == 2 and any(item.lower() in sender_folder for item in BNCC):
            base_path = os.path.join(docs_dir, "BNCC", sender_folder, subject_folder)
        elif selected_option == 3 and any(item.lower() in sender_folder for item in TECH):
            base_path = os.path.join(docs_dir, "TECH", sender_folder, subject_folder)
        
        if len(base_path) > MAX_PATH_LENGTH - 50:
            subject_folder = subject_folder[:MAX_PATH_LENGTH - len(docs_dir) - len(sender_folder) - 10]
            base_path = os.path.join(docs_dir, sender_folder, subject_folder)
        
        create_directory_if_not_exists(base_path)
        saved_file_path = join_without_overwriting(base_path, file_name=attachment.FileName)
        
        if len(saved_file_path) > MAX_PATH_LENGTH:
            max_filename_length = MAX_PATH_LENGTH - len(base_path) - len(extension) - 5
            filename = filename[:max_filename_length] + "_cut"
            saved_file_path = join_without_overwriting(base_path, file_name=filename + extension)

        try:
            attachment.SaveAsFile(saved_file_path)

            if openai_key is not None and extension.lower() in [".pdf", ".docx"]:
                try:
                    prompt = "Por favor, pegue esse conteúdo de texto abaixo e faça um resumo. "
                    prompt += "Use quebra de linhas para facilitar a leitura do usuário. "
                    prompt += "Daqui pra baixo, se não tiver nenhum conteúdo, "
                    prompt += "não faça um resumo, avise que não foi possível extrair o conteúdo do documento."
                    prompt += "\n\n"
                    content = ""
                    if extension.lower() == ".pdf":
                        content = read_pdf(saved_file_path)
                    elif extension.lower() == ".docx":
                        content = read_docx(saved_file_path)

                    if len(content) == 0:
                        raise Exception(f"sem conteúdo de texto no documento: {saved_file_path}.")

                    prompt += content
                    summary = send_prompt(openai_key, prompt)
                    print(summary)
                    summary_path = join_without_overwriting(base_path, file_name=filename[:len(filename) - 12] + "_resumo.txt")
                    with open(summary_path, "w", encoding="utf8") as f:
                        f.write(summary)
                except Exception as e:
                    print(f"Erro ao fazer resumo do arquivo '{attachment.FileName}': {e}")
        except Exception as e:
            print(f"Erro ao salvar '{attachment.FileName}': {e}")

    return cont > 0


def get_inbox():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    return inbox


def check_email(base_dir: str, openai_key: str | None = None) -> list[dict[str, str]]:
    """
    Varre emails não lidos na Caixa de Entrada e salva os documentos.
    :param openai_key: chave do gpt
    :param base_dir: pasta onde os documentos serão salvos
    :return: lista de dicionários contendo as principais informações das mensagens processadas
    """
    inbox = get_inbox()
    data = []
    current_year = datetime.now().year

    #Cria a tela para a escolha do filtro
    root = tk.Tk()
    root.title("Escolha uma opção")
    def clicked_button(value):
            root.destroy()
            global selected_option
            selected_option = value

    btn_no_filter = tk.Button(root, text="Sem filtro", command=lambda: clicked_button(1))
    btn_bncc = tk.Button(root, text="BNCC", command=lambda: clicked_button(2))
    btn_tech = tk.Button(root, text="TECH", command=lambda: clicked_button(3))

    btn_no_filter.pack(pady=5)
    btn_bncc.pack(pady=5)
    btn_tech.pack(pady=5)

    root.mainloop()
    for message in inbox.Items:
        try:
            email_year = message.ReceivedTime.year
            if email_year != current_year:
                continue

            if save_attachments(message, selected_option, base_dir, openai_key):
                message_data = get_email_info(message)
                data.append(message_data)
        except Exception as e:
            print(f"Erro ao processar e-mail: {e}")

    return data
