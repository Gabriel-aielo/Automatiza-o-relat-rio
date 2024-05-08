#------------------------------------------Importando bibliotecas------------------------------------------

import pandas as pd # análise de dados em Python que oferece estruturas de dados flexíveis e eficientes para trabalhar com dados tabulares.
import pytz # lidar com fusos horários em Python
import re # fornece suporte para expressões regulares em Python, permitindo realizar operações avançadas de busca e manipulação de strings.
from datetime import datetime # fornece classes para manipulação de datas e horários em Python.

#------------------------------------------Acessando os arquivos------------------------------------------
base_data_rendimentos = pd.read_csv('Rendimento.csv', sep=';')
base_data_descricao = pd.read_csv('descricaoDeItens.csv', sep=';', dtype={'Item': int, 'Ordem': float})
#------------------------------------------------------------------------------------------------------------------------------

#------------------------------------------Função para corrigir o formato das strings------------------------------------------
def corrigir_formato_porcentagem(string):
    # Verifica se a string contém pontos como separadores de milhares
    if re.search(r'\d+\.\d+\,\d+%', string):
        # Substitui os pontos por espaços vazios e a vírgula por um ponto
        string_corrigida = string.replace('.', '').replace(',', '.')
        return string_corrigida
    else:
        return string

#------------------------------------------Aplica a função à coluna 'Rendimentos %'------------------------------------------
base_data_rendimentos['Rendimentos %'] = base_data_rendimentos['Rendimentos %'].apply(corrigir_formato_porcentagem)

base_data_rendimentos['Rendimentos %'] = base_data_rendimentos['Rendimentos %'].str.replace('%', '').str.replace(',', '.').astype(float)

#------------------------------------------Definindo o fuso horário para Brasília------------------------------------------
tz = pytz.timezone('America/Sao_Paulo')

#------------------------------------------Filtrando os dados e aplicando a formatação condicional (rendimento finish)------------------------------------------
filtros_rendimento_finish_Tipo = base_data_rendimentos.loc[(base_data_rendimentos['TIPO'] == 'RENDIMENTO_1') & (base_data_rendimentos['Item'].astype(str).str.startswith('1'))].copy()

filtros_rendimento_finish_Tipo['TIPO'] = filtros_rendimento_finish_Tipo['TIPO'].replace('RENDIMENTO_1', 'Rendimento Finish')

filtros_rendimento_finish_Rendimento = filtros_rendimento_finish_Tipo

filtros_rendimento_finish_Tipo_Rendimento = filtros_rendimento_finish_Rendimento.loc[(filtros_rendimento_finish_Rendimento['Rendimentos %'] >= 120) | (filtros_rendimento_finish_Rendimento['Rendimentos %'] <= 80)].copy()
pd.set_option('display.max_columns', None)
#------------------------------------------------------------------------------------------------------------------------------

#------------------------------------------Filtrando os dados e aplicando a formatação condicional (rendimento bulk)------------------------------------------
filtros_rendimento_bulk_Tipo = base_data_rendimentos.loc[(base_data_rendimentos['TIPO'] == 'RENDIMENTO_1') & (base_data_rendimentos['Item'].astype(str).str.startswith('2'))]

filtros_rendimento_bulk_Tipo.loc[:, 'TIPO'] = filtros_rendimento_bulk_Tipo['TIPO'].replace('RENDIMENTO_1', 'Rendimento Bulk')

filtros_rendimento_bulk_Rendimento = filtros_rendimento_bulk_Tipo.copy()
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------

#------------------------------------OBTENDO A DESCRIÇÃO DOS ITENS------------------------------------
codigoItemListaDescricao = base_data_descricao['codigo item'].astype(str)
codigoItemBaseDataRendimentos = filtros_rendimento_finish_Tipo_Rendimento['Item'].astype(str)

print("Códigos de itens presentes nos rendimentos:")
print(codigoItemBaseDataRendimentos)

if codigoItemBaseDataRendimentos.isin(codigoItemListaDescricao).any():
    print("Pelo menos um código está presente na lista de descrição.")
    print("Códigos e descrições presentes na lista de descrição:")
    codigos_descricoes = base_data_descricao.loc[base_data_descricao['codigo item'].astype(str).isin(codigoItemBaseDataRendimentos)].copy()
    print("------------------------------------------------------------------------")
    print(codigos_descricoes[['codigo item']], codigos_descricoes[['descricao item']])
    print("------------------------------------------------------------------------")
#------------------------------------------------------------------------------------------------------------------------------------------------------

# Selecionando apenas as colunas desejadas, incluindo 'descricao item'
colunas_desejadas = ['Ordem', 'TIPO', 'Item', 'Rendimentos %']
filtros_rendimento_finish_Tipo_Rendimento = filtros_rendimento_finish_Tipo_Rendimento[colunas_desejadas]

# Verificando se o DataFrame não está vazio
if not filtros_rendimento_finish_Tipo_Rendimento.empty:
    # Convertendo a coluna 'codigo item' para int64
    codigos_descricoes['codigo item'] = codigos_descricoes['codigo item'].astype('int64')

    # Mesclando os DataFrames com base na coluna comum 'codigo item'
    merged_df = pd.merge(filtros_rendimento_finish_Tipo_Rendimento, codigos_descricoes, left_on='Item', right_on='codigo item', how='left')

    # Renomeando a coluna 'descricao item' para 'descricao'
    merged_df.rename(columns={'descricao item': 'Descrição'}, inplace=True)

    # Exibindo o DataFrame mesclado
    #print("DataFrame Mesclado:")
    #print(merged_df)

#------------------------------------------Salvando o DataFrame mesclado em um arquivo HTML------------------------------------------
    with open('filtros_rendimento_Finish.html', 'w', encoding="utf-8") as f:
        f.write('<!DOCTYPE html>\n')
        f.write('<html>\n')
        f.write('<head>\n')
        f.write('<meta charset="UTF-8">\n')
        f.write('<title>Filtros de Rendimento</title>\n')
        f.write('<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">\n')
        f.write('<style> .container { display: flex; justify-content: center; align-items: center; height: 100vh; } </style>\n')
        f.write('<style> .table-container { max-height: 80vh; overflow-y: auto; } </style>\n')
        f.write('<style> th, td { border: 1px solid #dddddd; text-align: left; padding: 8px; } </style>\n')
        f.write('<body>\n')
        f.write('<div class="container">')
        f.write('<div class="row justify-content-center text-center">')
        f.write('<h3 class="col-md-12">Relatório Finish</h3>')
        f.write('<div class="table-container">')

        f.write(merged_df.to_html(index=False, classes='table table-striped'))
        f.write('<small>Atualizado em {}</small>\n'.format(datetime.now(tz).strftime("%d/%m/%Y %H:%M:%S")))
        f.write('</div>')
        f.write('</div>')
        f.write('</div>')
        f.write('</body>\n')
        f.write('</html>\n')
else:
    with open('filtros_rendimento_Finish.html', 'w', encoding="utf-8") as f:
        f.write('<!DOCTYPE html>\n')
        f.write('<html>\n')
        f.write('<head>\n')
        f.write('<meta charset="UTF-8">\n')
        f.write('<title>Rendimento</title>\n')
        f.write('<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">\n')
        f.write('<style> .container { display: flex; justify-content: center; align-items: center; height: 100vh; } </style>\n')
        f.write('<style> .table-container { max-height: 80vh; overflow-y: auto; } </style>\n')
        f.write('</head>\n')
        f.write('<body>\n')
        f.write('<div class="container">')
        f.write('<div class="table-container">')
        f.write('<p>Nenhum caso crítico de Finish encontrado</p>\n')
        f.write('<small>Atualizado em {}</small>\n'.format(datetime.now(tz).strftime("%d/%m/%Y %H:%M:%S")))
        f.write('</div>')
        f.write('</div>')
        f.write('</body>\n')
        f.write('</html>\n')
