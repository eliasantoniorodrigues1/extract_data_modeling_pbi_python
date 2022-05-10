import os
import json
import pandas as pd
import re
from util import unzip_file


"""
    1 - No Power BI ir em Arquivo / Exportar / Modelo Power BI
    2 - Salvar o arquivo na pasta do script python: 
        \\path_to_your_file\\data_model
    3 - Colocar o nome do seu arquivo .pbit na linha 71 sem a extensão
    4 - Executar o script
"""
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data_model')


def extrai_comentario(string):
    regex = r"^(\/\/).*\n"

    try:
        if re.search(regex, string, re.MULTILINE):
            return re.search(regex, string, re.MULTILINE).group().replace('//', '').strip()
    except:
        return ''


def remove_comentario(string):
    regexp = re.compile(r"^(\/\/).*\n", re.MULTILINE)

    try:
        new_string = re.sub(regexp, '', string)
        if new_string:
            return new_string
    except:
        return string



def renomeia_arq(dash_name: str) -> None:
    for root, _, files in os.walk(os.path.join(DATA_DIR, dash_name)):
        for file in files:
            if 'DataModelSchema' in file:
                base_path, ext = os.path.splitext(os.path.join(root, file))

                if not ext:
                    os.rename(base_path, base_path + ".json")
                    print('Arquivo renomeado com sucesso!')


def gera_data_frame(dash_name: str) -> pd.DataFrame:
    with open(os.path.join(DATA_DIR, dash_name, 'DataModelSchema.json'), 'rb') as f:
        dict_data_model = json.load(f)

        frames = []
        for dado in dict_data_model['model']['tables']:
            if 'measures' in dado.keys():
                # faz uma varredura para achar medidas espalhadas pelo modelo
                frames.append(pd.DataFrame(dado['measures']))

    # agrega todas as listas de medidas em um só data frame
    df = pd.concat(frames)
    # gera coluna comentario
    df['Descrição'] = df['expression'].apply(extrai_comentario)
    df['Relatório'] = [dash_name for i in range(len(df))]
    df['expression'] = df['expression'].apply(remove_comentario)

    # renomeia as colunas do dash
    df.rename(columns={'expression': 'Fórmula',
              'name': 'Conceito'}, inplace=True)
    # reordena colunas
    df = df[['Relatório', 'Conceito', 'Descrição', 'Fórmula']]
    return df


if __name__ == '__main__':
    # nome do diretorio contendo o modelo de dados
    dash_name = 'RAC_Performance_Midia'
    # create directory to receive unzip files
    try:
        os.mkdir(os.path.join(DATA_DIR, dash_name))
    except Exception as e:
        print(f'Diretório já existe. {e}')
    # unzip pbit to get access to data files.
    unzip_file(os.path.join(DATA_DIR, dash_name + '.pbit'),
               os.path.join(DATA_DIR, dash_name))
    # adiciona a extensao .json no arquivo DataModelSchema
    try:
        renomeia_arq(dash_name=dash_name)
    except Exception as e:
        print(f'Arquivo já existe com esse nome. {e}')

    df = gera_data_frame(dash_name=dash_name)
    file_name = f'Dicionario_Dash_{dash_name}.xlsx'
    df.to_excel(file_name, index=False)
    print('Dados salvo com sucesso!')
