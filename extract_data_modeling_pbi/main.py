import os
import json
import pandas as pd


"""
    1 - No Power BI ir em Arquivo / Exportar / Modelo Power BI
    2 - Salvar o arquivo na pasta do seu projeto 03.HowTo
    3 - Abrir o arquivo com um descompactador zip, 7zip ou winrar
    4 - Extrair o arquivo DataModelSchema
    5 - Setar na variavel abaixo (path_json_file) o caminho do arquivo que voce
    acabou de extrair.
"""
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
folder = r'C:\Users\elias.rodrigues\Unidas\MKT_Geral - Documentos\03.Midia_Geral\01.Midia_em_Geral\01.RAC\11.Pre-Pagamento\03.HowTo'
file_name = 'DataModelSchema.json'

with open(os.path.join(folder, file_name), 'rb') as f:
    dict_data_model = json.load(f)
    df = pd.DataFrame(dict_data_model['model']
                      ['tables'][0]['measures'])

# print(dict_data_model['model']['tables'][0]['measures'])
print(f'Medidas salvas com sucesso em {BASE_DIR}')
print(df.head())
df.to_excel('Dicionario_Dash.xlsx', index=False)
