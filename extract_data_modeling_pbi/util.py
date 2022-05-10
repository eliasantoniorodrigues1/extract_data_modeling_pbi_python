import socket
import sys
import unicodedata
import re
import getpass
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
import json
import os
import time
import pandas as pd
import pathlib
import zipfile


def remove_acentos(string):
    new_string = str(string)
    normalizado = unicodedata.normalize('NFKD', new_string)

    return ''.join([c for c in normalizado if not unicodedata.combining(c)])


def remove_non_digit(string: str):
    new_str = str(string)
    return re.sub(r'\D', '', new_str)


def grava_log(caminho_completo, conteudo):
    with open(caminho_completo, 'w') as file:
        if conteudo != '':
            for linha in conteudo:
                file.write(linha)
                file.write('\n')
        else:
            file.write('Não há chapas para exclusao.')
            file.write('\n')


def dados_sessao_windows():
    hostname = socket.gethostname()
    local_ip = socket.gethostbyname(hostname)
    user = getpass.getuser()
    return hostname, local_ip, user


def cria_csv(nome, lista_cabecalho, conteudo):
    ...


def data_agora_str():
    return datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")


def envia_email(to, subject, messsage, filename=None, imagem=None, frm=None, host='mail.grupoaec.com.br', port=2525, password=None):

    LOG_DIR = os.path.join(BASE_DIR, 'log')

    msg = MIMEMultipart()
    msg['from'] = frm
    msg['to'] = to
    msg['cc'] = 'xxx@xxx.com.br'
    msg['subject'] = subject

    with open(os.path.join(LOG_DIR, filename), 'r') as file:
        attachment = MIMEText(file.read())
        attachment.add_header('Content-Disposition',
                              'attachment', filename=filename)

    corpo = MIMEText(messsage, 'html')
    msg.attach(corpo)

    # Tratando a imagem
    with open(imagem, 'rb') as img:
        email_image = MIMEImage(img.read())
        email_image.add_header('Content-ID', '<imagem>')
        msg.attach(email_image)

    with smtplib.SMTP(host=host, port=port) as smtp:
        try:
            smtp.ehlo()
            smtp.starttls()
            smtp.login(login, password)
            smtp.send_message(msg)
        except Exception as e:
            print('Email não enviado...')
            print('Erro:', e)


def salva_arquivo_json(arquivo, dados):
    with open(arquivo, 'w', encoding='utf-8') as f:
        json.dump(dados, f, ensure_ascii=False, indent=4, sort_keys=False)


def limpa_clipboard():
    if sys.platform == 'win32':
        import win32clipboard as clip
        clip.OpenClipboard()
        clip.EmptyClipboard()
        clip.CloseClipboard()
    elif sys.platform.startswith('linux'):
        import subprocess
        proc = subprocess.Popen(('xsel', '-i', '-b', '-1', '/dev/null'),
                                stdin=subprocess.PIPE)
        proc.stdin.close()
        proc.wait()
    else:
        raise RuntimeError(
            'Plataforma não suportada para limpar memória de items copiados.')


def le_json(arquivo: str):
    with open(arquivo, encoding='utf-8') as arquivo:
        dados = json.load(arquivo)

    return dados


def atualiza_json(arq: str, chave: str, lista: list, indice: int):
    with open(arq, 'r', encoding='utf-8') as file:
        data = json.load(file)

    data[indice][chave] = lista

    with open(arq, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4)


def quartil(lista: list) -> list:
    amostragem = sorted(lista)
    indice_principal = int(len(amostragem)/2)
    # Declara lista Inferior, geral e superior para calcular o quartil
    consolidado = [amostragem[: indice_principal],
                   amostragem, amostragem[indice_principal:]]

    print(consolidado)
    dicionario = {}
    for idx, lista in enumerate(consolidado):
        i = int(len(lista)/2)
        if len(lista) % 2 == 0:
            q = (lista[i-1] + lista[i]) / 2
        else:
            q = lista[i]

        dicionario[f'{idx + 1}º Quartil'] = q

    return dicionario


def validate_email(email: str) -> str:
    new_email = str(email).lower()
    # Regex official RFC 5322 Official Standard)
    regexp = re.compile(r"(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*|(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*)@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])")
    # A expressão oficial não trata emials com acento, por isso eu removo na função
    # abaixo
    new_email = remove_acentos(new_email)
    retorno = re.match(regexp, new_email)
    if retorno:
        return 'Email válido'
    else:
        return 'Email inválido'


def validate_phone(phone_number: str) -> str:
    regexp = re.compile(
        r"^|\()?\s*(\d{2})\s*(\s|\))*(9?\d{4})(\s|-)?(\d{4})($|\n)")
    retorno = re.match(regexp, phone_number)
    if retorno:
        return 'Telefone válido'
    else:
        return 'Telefone inválido'


def consolida_bases(caminho: str):
    """
        params: caminho: Recebe um caminho para ler todos os arquivos csv ou xlsx 
        dentro da pasta.
        Para que a consolidação funcione os dados devem ter o mesmo número de colunas
        e nome.
    """
    # Consolida arquivos CSV
    for _, _, files in os.walk(caminho):
        # Consolida CSV
        dados_csv = pd.concat(
            # Criar função para detectar encoding do arquivo:
            [pd.read_csv(os.path.join(caminho, file), delimiter=';', encoding='latin-1') for file in files if file.endswith('.csv')])

        # Consolida Excel
        dados_excel = pd.concat(
            [pd.read_excel(os.path.join(caminho, file)) for file in files if file.endswith('.xls')])

    return dados_csv, dados_excel


def remove_duplicados(dataframe, coluna, manter, nome):
    """
        params: dataframe:
        params: coluna:
        params: manter:
        params: nome:
    """
    # Remove duplicados
    print(coluna, manter)
    df = dataframe.drop_duplicates(subset=coluna, keep=manter)

    # Salva dataset sem duplicidade:
    df.to_csv(f'{nome}.csv', sep=';')


def join_data_frames(df_left, df_right, column_left: str, column_right: str):
    """
        params: df_left: Dataframe principal onde eu quero consolidar as 
        informações.
        params: df_right: Dataframe secundário onde eu quero buscar uma
        coluna para meu dataframe principal.
        params: column_left: coluna do dataframe da esquerda
        params: column_right: coluna do dataframe da esquerda
    """
    df_joined = pd.merge(
        df_left, df_right, left_on=column_left, right_on=column_right)
    return df_joined


def meta_data_files(basedir: str, extension: str) -> list:
    """
        Recebe um diretório base e uma extensão para retornar um dicionário
        contendo a data de modificação do arquivo
        param: basedir: Diretório base de onde os arquivos estão localizado
        param: extension: extensão específica para ser buscada dentro do diretório
        afim de não atrapalhar os demais arquivos.
        Retorna apenas o atributo data de modiciação.

    """
    dict_dados = {}
    for path in pathlib.Path(basedir).iterdir():
        info = path.stat()
        mtime = info.st_mtime
        name, ext_file = os.path.splitext(path)
        name = name.split('\\')[-1]
        if ext_file == extension:
            dict_dados[name] = mtime  # Data de modificação

    return dict_dados


def max_dict_data(dict: dict):
    max_value = max(dict, key=dict.get)
    return max_value


def download_wait(path_to_downloads):
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < 20:
        time.sleep(1)
        dl_wait = False
        for fname in os.listdir(path_to_downloads):
            if fname.endswith('.crdownload'):
                dl_wait = True
        seconds += 1
    return seconds


def coleta_dados_email(sendermail: str, subject: str, date: str):
    '''
        Essa função faz uma busca no Outlook usando a biblioteca win32com da
        Microsoft, por parâmetros de: Subject, SenderMail, Date com objetivo
        de baixar um anexo específico ou vários.
        params: subject: assunto do email
        params: sendermail: email de quem fez o envio
        params: date: data especifica para buscar uma conta.
    '''


def unzip_file(path_to_zip_file: str, directory_to_extract_to: str) -> None:
    with zipfile.ZipFile(path_to_zip_file, 'r') as zip_ref:
        zip_ref.extractall(directory_to_extract_to)
        print(f'Dados extraídos para {directory_to_extract_to}')


if __name__ == '__main__':
    print(validate_email('josé@gmail.com.br'))
