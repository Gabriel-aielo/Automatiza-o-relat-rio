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
base_data_rendimentos['Rendimentos %'] = base_data_rendimentos['Rendimentos %'].astype(str).apply(corrigir_formato_porcentagem)

base_data_rendimentos['Rendimentos %'] = base_data_rendimentos['Rendimentos %'].str.replace('%', '').str.replace(',', '.').astype(float)

#------------------------------------------Definindo o fuso horário para Brasília------------------------------------------
tz = pytz.timezone('America/Sao_Paulo')

#------------------------------------------Código para filtrar os dados e aplicar a formatação condicional------------------------------------------
filtros_rendimento_bulk_Tipo = base_data_rendimentos.loc[(base_data_rendimentos['TIPO'] == 'RENDIMENTO_1') & (base_data_rendimentos['Item'].astype(str).str.startswith('2'))]

filtros_rendimento_bulk_Tipo.loc[:, 'TIPO'] = filtros_rendimento_bulk_Tipo['TIPO'].replace('RENDIMENTO_1', 'Rendimento Bulk')

filtros_rendimento_bulk_Rendimento = filtros_rendimento_bulk_Tipo.copy()

filtros_rendimento_bulk_Rendimento = filtros_rendimento_bulk_Rendimento.loc[filtros_rendimento_bulk_Rendimento['Rendimentos %'] != 100].copy()
pd.set_option('display.max_columns', None)

#------------------------------------------Selecionando apenas as colunas desejadas------------------------------------------
colunas_desejadas_bulk = ['Ordem', 'TIPO', 'Item','Rendimentos %']
filtros_rendimento_bulk_Rendimento = filtros_rendimento_bulk_Rendimento[colunas_desejadas_bulk].copy()

#------------------------------------------Verificando se o DataFrame está vazio------------------------------------------
if filtros_rendimento_bulk_Rendimento.empty:
    with open('filtros_rendimento_Bulk.html', 'w', ) as f:
        f.write('<!DOCTYPE html>\n')
        f.write('<html>\n')
        f.write('<head>\n')
        f.write('<meta charset="UTF-8">\n')
        f.write('<title>Rendimento Bulk</title>\n')
        f.write('<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">\n')
        f.write('<style> .container { display: flex; justify-content: center; align-items: center; height: 100vh; } </style>\n')
        f.write('<style> .table-container { max-height: 80vh; overflow-y: auto; } </style>\n')
        f.write('</head>\n')
        f.write('<body>\n')
        f.write('<div class="container">')
        f.write('<div class="table-container">')
        f.write('<p>Nenhum caso crítico de Bulk encontrado</p>\n')
        f.write('<small>Atualizado em {}</small>\n'.format(datetime.now(tz).strftime("%d/%m/%Y %H:%M:%S")))
        f.write('</div>')
        f.write('</div>')
        f.write('</body>\n')
        f.write('</html>\n')
else:
    # Carregando as descrições dos itens
    codigos_descricoes_bulk = base_data_descricao.loc[base_data_descricao['codigo item'].astype(str).isin(filtros_rendimento_bulk_Rendimento['Item'].astype(str))].copy()

    # Convertendo a coluna 'codigo item' para int64
    codigos_descricoes_bulk['codigo item'] = codigos_descricoes_bulk['codigo item'].astype('int64')

    # Mesclando os DataFrames com base na coluna comum 'codigo item'
    merged_df_bulk = pd.merge(filtros_rendimento_bulk_Rendimento, codigos_descricoes_bulk, left_on='Item', right_on='codigo item', how='left')

    # Renomeando a coluna 'descricao item' para 'Descrição'
    merged_df_bulk.rename(columns={'descricao item': 'Descrição'}, inplace=True)

    # Definindo a função para aplicar a formatação condicional
    def highlight_percentage_row_bulk(row):
        if row['Rendimentos %'] >= 120 or row['Rendimentos %'] <= 80:
            return ['background-color: rgba(255, 128, 128, 0.5); border: 1px solid black'] * len(row)
        else:
            return ['border: 1px solid black'] * len(row)

    # Aplicando a função de formatação condicional ao DataFrame
    styled_filtros_rendimento_bulk_Rendimento = merged_df_bulk.style.apply(highlight_percentage_row_bulk, axis=1)

    # Adicionando bordas para as colunas
    styled_filtros_rendimento_bulk_Rendimento.set_table_styles([{'selector': 'th', 'props': [('border', '1px solid black')]}])

    #------------------------------------------Salvando o DataFrame estilizado em um arquivo HTML------------------------------------------
    with open('filtros_rendimento_Bulk.html', 'w', encoding="utf-8") as f:
        f.write('<!DOCTYPE html>\n')
        f.write('<html>\n')
        f.write('<head>\n')
        f.write('<meta charset="UTF-8">\n')
        f.write('<title>Filtros de Rendimento Bulk</title>\n')
        f.write('<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">\n')
        f.write('<style> .container { display: flex; justify-content: center; align-items: center; height: 100vh; } </style>\n')
        f.write('<style> .table-container { max-height: 80vh; overflow-y: auto; } </style>\n')
        f.write('<style> th, td { border: 1px solid #dddddd; text-align: left; padding: 8px; } </style>\n')
        f.write('</head>\n')
        f.write('<body>\n')
        f.write('<div class="container">')
        f.write('<div class="row justify-content-center text-center">')
        f.write('<h3 class="col-md-12">Relatório Bulk</h3>')
        f.write('<div class="table-container">')
        f.write(styled_filtros_rendimento_bulk_Rendimento.to_html(index=False, classes='table table-striped'))
        f.write('<small>Atualizado em {}</small>\n'.format(datetime.now(tz).strftime("%d/%m/%Y %H:%M:%S")))
        f.write('</div>')
        f.write('</div>')
        f.write('</div>')
        f.write('</body>\n')
        f.write('</html>\n')

        print('Este são os rendimentos: ')
        print(filtros_rendimento_bulk_Rendimento.count())
