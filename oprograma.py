from pathlib import Path
from dataclasses import dataclass

import openpyxl

from selenium import webdriver
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions


SEARCH_BOX_XPATH = '//div[@data-testid="chat-list-search"]'
SEND_BUTTON_XPATH = '//button[@data-testid="compose-btn-send"]'


MESSAGE_FORMAT = """
    Olá, {nome}! Esta é uma mensagem de teste sendo enviada pelo programa da
    UP. Se você recebeu esta mensagem, então quer dizer que o programa funciona
    :)
"""


def connect_to_wpp(user_data_dir: Path = Path('.') / 'chrome-data') -> WebDriver:
    _ = user_data_dir
    driver = webdriver.Firefox()
    driver.get("https://web.whatsapp.com")
    return driver


def wait_for_qrcode_scan(driver: WebDriver) -> None:
    print('Esperando login/carregamento da página...')
    WebDriverWait(driver, 300).until(
        expected_conditions.presence_of_element_located((By.XPATH, SEARCH_BOX_XPATH))
    )
    print('Feito!')


def send_message(driver: WebDriver, to: str, msg: str) -> None:
    driver.get(f'https://web.whatsapp.com/send?phone={to}&text={msg}')

    print('Esperando botão de envio aparecer...')
    WebDriverWait(driver, 300).until(
        expected_conditions.presence_of_element_located((By.XPATH, SEARCH_BOX_XPATH))
    )
    import time
    time.sleep(5)
    print('Apareceu! Hora de clicar no botão...')
    send_button = driver.find_element(By.XPATH, value=SEND_BUTTON_XPATH)

    send_button.click()
    print('Clicou!')


@dataclass
class Contact:
    name: str
    state: str
    number: str  # Somente números
    contacted: bool


def normalize_number(number: str) -> str:
    import re

    numbers_only = ''.join(re.findall(r'\d', number))

    # Falta só o +55 e já tem o 9 extra
    if len(numbers_only) == len('048999999999'):
        return f"55{numbers_only}"

    # Falta só o +55 e não tem o 9 extra
    if len(numbers_only) == len('04899999999') and numbers_only[0] == '0':
        return f"55{numbers_only}"

    # Falta só o +55 0 e já tem o 9 extra
    if len(numbers_only) == len('48999999999'):
        return f"550{numbers_only}"

    # Falta só o +55 0 e já tem o 9 extra
    if len(numbers_only) == len('4899999999'):
        return f"550{numbers_only}"

    # Já tem o 9 extra:

    # Tem tudo:
    return numbers_only


def load_contacts(sheet_path: Path) -> list[Contact]:
    contacts = []
    workbook = openpyxl.load_workbook(sheet_path)

    _, *rows = workbook.active.iter_rows(min_row=2, values_only=True)

    for line in rows:
        if all(x is None for x in line):
            break

        _, state, name, number, contacted, *_ = line

        contacts.append(Contact(
            name=name,
            state=state,
            number=normalize_number(str(number)),
            contacted=contacted == '=TRUE()',
        ))

    return contacts


def make_message(contact: Contact) -> str:
    return ' '.join([s.strip() for s in MESSAGE_FORMAT.split('\n')]).format(
        nome=contact.name.split()[0],
    ).strip()


if __name__ == '__main__':
    import sys

    contacts = load_contacts(Path('CONTATOS UNIDADE POPULAR -SC.xlsx'))

    if len(sys.argv) > 1 and sys.argv[1] == '--ver-contatos':
        from pprint import pprint
        pprint(contacts)
        for contact in contacts:
            msg = make_message(contact)
            print(f"{msg=}")
        exit(0)

    with connect_to_wpp() as driver:
        wait_for_qrcode_scan(driver)

        print(f"{driver.get_cookies()=}")

        for contact in contacts:
            # to = '5504896870888'
            if contact.contacted:
                print(f"Pulando {contact.name} (já foi entrado em contato)")
                continue

            print(f'Enviando mensagem para {contact.name} ({contact.number})')
            msg = make_message(contact)
            send_message(driver, to=contact.number, msg=msg)
