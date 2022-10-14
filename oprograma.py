import time
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
Bom dia, camarada!

Estou entrando em contato com você em nome da Unidade Popular Pelo Socialismo
de Santa Catarina!

Você preencheu o formulário de pré-filiação e por isso estamos entrando em
contato para te convidar pra uma primeira atividade!

No sábado (15/10) às 17:30 vamos realizar uma Plenária de apresentação da UP.
Para além disso, vamos debater sobre nossa atuação nesse segundo turno das
eleições e a postura do nosso partido para combatermos o avanço do fascismo no
nosso país.

Essa plenária vai acontecer de forma híbrida para aqueles que não moram em
Florianópolis possam participar também! Caso possa participar nos avise que te
passamos melhor as informações!
"""

CONTACTED_LIST_FILE = Path("lista_contactados.txt")


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

    # Já tem o +55 e já tem o 9 extra mas não tem o 0
    if len(numbers_only) == len('5548999999999'):
        return f"55{numbers_only.lstrip('55')}"

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

    # Tem tudo:
    return numbers_only


def load_contacts(sheet_path: Path) -> list[Contact]:
    contacts = []
    workbook = openpyxl.load_workbook(sheet_path)

    _, *rows = workbook.active.iter_rows(min_row=2, values_only=True)

    for line in rows:
        if all(x is None for x in line):
            break

        _, state, name, number, contacted, _, email, *_ = line

        if (state is None
            or name is None
            or number is None
            or (email is not None
                and email.lower().strip() in ['repetido', 's/whats'])
        ):
            print(f"Pulando {name, number, email}")
            continue

        norm_number = normalize_number(str(number).removesuffix('.0'))
        number = norm_number

        for c in contacts:
            if name == c.name or number == c.number:
                print(
                    f"Pulando {name} "
                    f"(repetido: {c.name} = {name}, {c.number} = {number})"
                )
                continue

        contacts.append(Contact(
            name=name,
            state=state,
            number=number,
            contacted=contacted == '=TRUE()' if not isinstance(contacted, bool) else contacted,
        ))

    return contacts


def make_message(contact: Contact) -> str:
    return ' '.join([s.strip() if s else '\n\n' for s in MESSAGE_FORMAT.split('\n')]).format(
        nome=contact.name.split()[0],
    ).strip().replace('\n', '%0a')


if __name__ == '__main__':
    import sys

    contacts = load_contacts(Path('CONTATOS UNIDADE POPULAR -SC.xlsx'))

    start = 0
    for arg in sys.argv:
        try:
            start = int(arg)
            break
        except:
            pass

    not_contacted = [c for c in contacts[start:] if not c.contacted]
    print(f"Primeiro não contactado {not_contacted[0].name} (i = {contacts.index(not_contacted[0])})")

    if '--ver-contatos' in sys.argv:
        for i, contact in enumerate(contacts[start:], start=start):
            if contact.contacted:
                print(f"-- {i}: Pulando {contact.name} (já foi entrado em contato)")
                continue

            print(f'== {i} Enviando mensagem para {contact.name} ({contact.number})')
        exit(0)

    not_contacted = [c for c in contacts[start:] if not c.contacted]
    start = contacts.index(not_contacted[0])

    with connect_to_wpp() as driver:
        wait_for_qrcode_scan(driver)

        print(f"{driver.get_cookies()=}")

        x = 0
        for i, contact in enumerate(contacts[start:], start=start):
            if x == 10:
                print("Esperando um pouco enviar as que já foram...")
                time.sleep(10)

            if contact.contacted:
                print(f"-- {i}: Pulando {contact.name} (já foi entrado em contato)")
                continue

            x += 1

            print(f'== {i} Enviando mensagem para {contact.name} ({contact.number})')
            msg = make_message(contact)
            try:
                send_message(driver, to=contact.number, msg=msg)
                time.sleep(2)
                result = "Contactado"
            except:
                print(f"Não deu para mandar mensagem para {contact.name}")
                result = "Deu erro"

            with open(CONTACTED_LIST_FILE, "a") as f:
                f.write(f"\n{contact.name}, {contact.number}, {contact.state} ({result})")
