import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
import plotly.express as px
from datetime import datetime
from tabulate import tabulate
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from IPython.display import HTML  # Adicionando a importação necessária
from io import BytesIO
import base64

base_Data_Bruto = pd.read_csv('CS0530&(CSV).csv', sep=';')

# Remover linhas que têm mais de 5 NaNs
base_data = base_Data_Bruto.dropna(thresh=len(base_Data_Bruto.columns) - 5)

#----------------------------------------------------------------------------------------------------------

# Filtrando e contando
contagemACAsemREQ = base_data['ACA sem REQ'].eq("Sim").sum()
filtroLinhaACAsemREQ = base_data.loc[base_data['ACA sem REQ'] == "Sim"]

contagemREQsemACA = (base_data[(base_data['REQ sem ACA'] == "Sim") & (base_data['Estado'] == "Finalizada")])['REQ sem ACA'].count()
filtroLinhaREQsemACA = base_data.loc[(base_data['REQ sem ACA'] == "Sim") & (base_data['Estado'] == "Finalizada")]

contagemACAsemGGF = base_data['ACA sem GGF'].eq("Sim").sum()
filtroLinhaACAsemGGF = base_data.loc[base_data['ACA sem GGF'] == "Sim"]

contagemGGFsemACA = base_data['GGF sem ACA'].eq("Sim").sum()
filtroLinhaGGFsemACA = base_data.loc[base_data['GGF sem ACA'] == "Sim"]

#contagemACAsemMOB = base_data['ACA sem MOB'].eq("Não").sum()
#filtroLinhaACAsemMOB = base_data.loc[base_data['ACA sem MOB'] == "Não"]
contagemACAsemMOB = base_data[(base_data['ACA sem MOB'] == "Sim") & (base_data['Tipo'] == "Finish")]['ACA sem MOB'].count()
filtroLinhaACAsemMOB = base_data.loc[(base_data['ACA sem MOB'] == "Sim") & (base_data['Tipo'] == "Finish")]

contagemMOBsemACA = base_data['MOB sem ACA'].eq("Sim").sum()
filtroLinhaMOBsemACA = base_data.loc[base_data['MOB sem ACA'] == "Sim"]

contagemRRQSemREQItemcor = base_data['RRQ Sem REQ Item cor.'].eq("Sim").sum()
filtroLinhaRRQSemREQ = base_data.loc[base_data['RRQ Sem REQ Item cor.'] == "Sim"]

contagemRestriçãoGGFMOB = base_data['Restrição GGF/MOB'].eq("Sim").sum()
filtroLinhaGGFMOB = base_data.loc[base_data['Restrição GGF/MOB'] == "Sim"]

#Visualização da data mais recente, para indicar quando teve uma última movimentação.
data_mais_recente = base_data['Últ Movto'].max()

print(data_mais_recente)

print('Contagem ACA sem REQ:')
print(contagemACAsemREQ)

print('------------------------------')

print('Contagem REQ sem ACA')
print(contagemREQsemACA)

print('------------------------------')

print('Contagem ACA sem GGF')
print(contagemACAsemGGF)

print('------------------------------')

print('Contagem GGF sem ACA')
print(contagemGGFsemACA)

print('------------------------------')

print('Contagem ACA sem MOB')
print(contagemACAsemMOB)

print('------------------------------')

print('Contagem MOB sem ACA')
print(contagemMOBsemACA)

print('------------------------------')

print('Contagem RRQ sem REQ')
print(contagemRRQSemREQItemcor)

print('------------------------------')

print('Contagem Restrição GGF MOB')
print(contagemRestriçãoGGFMOB)

print('---------------------Apresentação resultados---------------------')
if contagemACAsemREQ != 0:
    print('Contagem ACA sem REQ:')
    print(contagemACAsemREQ)
    print('Linhas com ACA sem REQ:')
    print(tabulate(filtroLinhaACAsemREQ, headers='keys', tablefmt='pretty'))
    print('------------------------------------------')

if contagemREQsemACA != 0:
    print('Contagem REQ sem ACA')
    print(contagemREQsemACA)
    print('Linhas com REQ sem ACA:')
    print(tabulate(filtroLinhaREQsemACA, headers='keys', tablefmt='pretty'))
    print('------------------------------------------')

if contagemACAsemGGF != 0:
    print('Contagem ACA sem GGF')
    print(contagemACAsemGGF)
    print('Linhas com ACA sem GGF:')
    print(tabulate(filtroLinhaACAsemGGF, headers='keys', tablefmt='pretty'))
    print('------------------------------------------')

if contagemGGFsemACA != 0:
    print('Contagem GGF sem ACA')
    print(contagemGGFsemACA)
    print('Linhas com GGF sem ACA:')
    print(tabulate(filtroLinhaGGFsemACA, headers='keys', tablefmt='pretty'))
    print('------------------------------------------')

if contagemACAsemMOB != 0:
    print('Contagem ACA sem MOB')
    print(contagemACAsemMOB)
    print('Linhas com ACA sem MOB:')
    print(tabulate(filtroLinhaACAsemMOB, headers='keys', tablefmt='pretty'))
    print('------------------------------------------')

if contagemMOBsemACA != 0:
    print('Contagem MOB sem ACA')
    print(contagemMOBsemACA)
    print('Linhas com MOB sem ACA:')
    print(tabulate(filtroLinhaMOBsemACA, headers='keys', tablefmt='pretty'))
    print('------------------------------------------')

if contagemRRQSemREQItemcor != 0:
    print('Contagem RRQ sem REQ')
    print(contagemRRQSemREQItemcor)
    print('Linhas com RRQ sem REQ:')
    print(tabulate(filtroLinhaRRQSemREQ, headers='keys', tablefmt='pretty'))
    print('------------------------------------------')

if contagemRestriçãoGGFMOB != 0:
    print('Contagem Restrição GGF MOB')
    print(contagemRestriçãoGGFMOB)
    print('Linhas com Restrição GGF MOB:')
    print(tabulate(filtroLinhaGGFMOB, headers='keys', tablefmt='pretty'))
    print('------------------------------------------')

#--------------------------------------------------------------------------------------------------------

# Definindo os dados
categorias = ['ACA sem REQ', 'REQ sem ACA', 'ACA sem GGF', 'GGF sem ACA', 'ACA sem MOB', 'MOB sem ACA', 'RRQ sem REQ', 'Restrição GGF MOB']
contagens = [contagemACAsemREQ, contagemREQsemACA, contagemACAsemGGF, contagemGGFsemACA, contagemACAsemMOB, contagemMOBsemACA, contagemRRQSemREQItemcor, contagemRestriçãoGGFMOB]

# Arredondar os valores das contagens para números inteiros
contagens_arredondadas = [int(round(contagem)) for contagem in contagens]

# Criar o gráfico de barras
plt.figure(figsize=(10, 6))
plt.barh(categorias, contagens_arredondadas, color='skyblue')
plt.xlabel('Contagem')
plt.ylabel('Categoria')
plt.title('Contagem de Problemas')
plt.gca().invert_yaxis()  # Inverte a ordem das categorias no eixo y

# Salvar o gráfico em formato de imagem antes de exibi-lo
plt.savefig('grafico.png')

# Exibir o gráfico no Google Colab
plt.show()

# Salvar o gráfico em formato de imagem
with open('grafico.png', 'rb') as f:
    image_base64 = base64.b64encode(f.read()).decode('utf-8')

#---------------------------------------------------------------------------------------------------------#-------------------------------------------------------------------------------------------------

# Criar um novo arquivo HTML
arquivoHTML = "ORDENS_CRITICAS.html"

# Escrever o cabeçalho do arquivo HTML e incluir o Bootstrap CSS
with open(arquivoHTML, "w", encoding="utf-8") as f:
    f.write('<html lang="pt-BR">\n<head>\n<title>Resultados dos Filtros</title>\n')
    f.write('<meta charset="UTF-8">\n')
    f.write('<link href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" rel="stylesheet">\n')
    f.write('<style>')
    f.write('.table-scrollable { max-width: 100%; overflow-x: auto; }')
    f.write('.center-content { text-align: center; }')
    f.write('.no-wrap td { white-space: nowrap; }')  # Aplica a classe no-wrap às células da tabela
    f.write('.table td, .table th { padding: 0.75rem; }')  # Ajusta o padding das células
    f.write('p{margin: 0;}')
    f.write('table{font-size:10px}')
    f.write('</style>')
    f.write("</head>\n<body>\n")

    # Escrever o gráfico no arquivo HTML
    f.write("<div class='container-fluid font-weight-bolder text-center bg-light mb-5 p-5 shadow-sm'>\n")
    f.write("<h4>RELATÓRIO DE CUSTOS</h4>")
    f.write("</div>")

    # Adicionar os resultados dos filtros ao arquivo HTML
    f.write("<div class='container p-5 rounded shadow'>\n")
    f.write("<div class='row'>\n")
    f.write("<h3>ORDENS CRÍTICAS:<br></h3>")

    f.write('<div class="col-md-12 table table-bordered center-content no-wrap table-responsive">')

    if contagemACAsemREQ != 0:
        f.write('<div class="col-md-12">')
        f.write(f'<h4>Casos ACA sem REQ:{contagemACAsemREQ}</h4>')
        f.write('<table class="table table-bordered center-content no-wrap">')
        # Aplica a classe no-wrap aos cabeçalhos da tabela
        f.write(filtroLinhaACAsemREQ[['Item', 'Descrição', 'Nr Ordem', 'Data Emissão','Estado', 'Qt Produzida', 'ACA sem REQ', 'REQ sem ACA', 'ACA sem GGF', 'GGF sem ACA', 'ACA sem MOB', 'MOB sem ACA', 'RRQ Sem REQ Item cor.', 'Restrição GGF/MOB', 'Últ Movto', 'Tipo']].to_html(index=False, classes="no-wrap"))
        f.write('</table>')
        f.write('</div>')
        f.write('<hr>')
        f.write('<br>')

    if contagemREQsemACA != 0:
        f.write('<div class="col-md-12">')
        f.write(f'<h4>Casos REQ sem ACA:{contagemREQsemACA}</h4>')
        f.write('<table class="table table-bordered center-content no-wrap">')
        # Aplica a classe no-wrap aos cabeçalhos da tabela
        f.write(filtroLinhaREQsemACA[['Item', 'Descrição', 'Nr Ordem', 'Data Emissão', 'Estado', 'Qt Produzida', 'ACA sem REQ', 'REQ sem ACA', 'ACA sem GGF', 'GGF sem ACA', 'ACA sem MOB', 'MOB sem ACA', 'RRQ Sem REQ Item cor.', 'Restrição GGF/MOB', 'Últ Movto', 'Tipo']].to_html(index=False, classes="no-wrap"))
        f.write('</table>')
        f.write('</div>')
        f.write('<hr>')
        f.write('<br>')

    if contagemACAsemGGF != 0:
        f.write('<div class="col-md-12">')
        f.write(f'<h4>Casos ACA sem GGF:{contagemACAsemGGF}</h4>')
        f.write('<table class="table table-bordered center-content no-wrap">')
        # Aplica a classe no-wrap aos cabeçalhos da tabela
        f.write(filtroLinhaACAsemGGF[['Item', 'Descrição', 'Nr Ordem', 'Data Emissão', 'Estado', 'Qt Produzida', 'ACA sem REQ', 'REQ sem ACA', 'ACA sem GGF', 'GGF sem ACA', 'ACA sem MOB', 'MOB sem ACA', 'RRQ Sem REQ Item cor.', 'Restrição GGF/MOB', 'Últ Movto', 'Tipo']].to_html(index=False, classes="no-wrap"))
        f.write('</table>')
        f.write('</div>')
        f.write('<hr>')
        f.write('<br>')

    if contagemGGFsemACA != 0:
        f.write('<div class="col-md-12">')
        f.write(f'<h4>Casos GGF sem ACA: {contagemGGFsemACA}</h4>')
        f.write('<table class="table table-bordered center-content no-wrap">')
        # Aplica a classe no-wrap aos cabeçalhos da tabela
        f.write(filtroLinhaGGFsemACA[['Item', 'Descrição', 'Nr Ordem', 'Data Emissão', 'Estado', 'Qt Produzida', 'ACA sem REQ', 'REQ sem ACA', 'ACA sem GGF', 'GGF sem ACA', 'ACA sem MOB', 'MOB sem ACA', 'RRQ Sem REQ Item cor.', 'Restrição GGF/MOB', 'Últ Movto', 'Tipo']].to_html(index=False, classes="no-wrap"))
        f.write('</table>')
        f.write('</div>')
        f.write('<hr>')
        f.write('<br>')

    if contagemACAsemMOB != 0:
        f.write('<div class="col-md-12">')
        f.write(f'<h4>Casos ACA sem MOB:{contagemACAsemMOB}</h4>')
        f.write('<table class="table table-bordered center-content no-wrap">')
        # Aplica a classe no-wrap aos cabeçalhos da tabela
        f.write(filtroLinhaACAsemMOB[['Item', 'Descrição', 'Nr Ordem', 'Data Emissão','Estado', 'Qt Produzida', 'ACA sem REQ', 'REQ sem ACA', 'ACA sem GGF', 'GGF sem ACA', 'ACA sem MOB', 'MOB sem ACA', 'RRQ Sem REQ Item cor.', 'Restrição GGF/MOB', 'Últ Movto', 'Tipo']].to_html(index=False, classes="no-wrap"))
        f.write('</table>')
        f.write('</div>')
        f.write('<hr>')
        f.write('<br>')

    if contagemMOBsemACA != 0:
        f.write('<div class="col-md-12">')
        f.write(f'<h4>Casos MOB sem ACA:{contagemMOBsemACA}</h4>')
        f.write('<table class="table table-bordered center-content no-wrap">')
        # Aplica a classe no-wrap aos cabeçalhos da tabela
        f.write(filtroLinhaMOBsemACA[['Item', 'Descrição', 'Nr Ordem', 'Data Emissão','Estado', 'Qt Produzida', 'ACA sem REQ', 'REQ sem ACA', 'ACA sem GGF', 'GGF sem ACA', 'ACA sem MOB', 'MOB sem ACA', 'RRQ Sem REQ Item cor.', 'Restrição GGF/MOB', 'Últ Movto', 'Tipo']].to_html(index=False, classes="no-wrap"))
        f.write('</table>')
        f.write('</div>')
        f.write('<hr>')
        f.write('<br>')

    if contagemRRQSemREQItemcor != 0:
        f.write('<div class="col-md-12">')
        f.write(f'<h4>Casos RRQ sem REQ:{contagemRRQSemREQItemcor}</h4>')
        f.write('<table class="table table-bordered center-content no-wrap">')
        # Aplica a classe no-wrap aos cabeçalhos da tabela
        f.write(filtroLinhaRRQSemREQ[['Item', 'Descrição', 'Nr Ordem', 'Data Emissão','Estado', 'Qt Produzida', 'ACA sem REQ', 'REQ sem ACA', 'ACA sem GGF', 'GGF sem ACA', 'ACA sem MOB', 'MOB sem ACA', 'RRQ Sem REQ Item cor.', 'Restrição GGF/MOB', 'Últ Movto', 'Tipo']].to_html(index=False, classes="no-wrap"))
        f.write('</table>')
        f.write('</div>')
        f.write('<hr>')
        f.write('<br>')

    if contagemRestriçãoGGFMOB != 0:
        f.write('<div class="col-md-12">')
        f.write(f'<h4>Casos Restrição GGF MOB:{contagemRestriçãoGGFMOB}</h4>')
        f.write('<table class="table table-bordered center-content no-wrap">')
        # Aplica a classe no-wrap aos cabeçalhos da tabela
        f.write(filtroLinhaGGFMOB[['Item', 'Descrição', 'Nr Ordem', 'Data Emissão', 'Estado', 'Qt Produzida', 'ACA sem REQ', 'REQ sem ACA', 'ACA sem GGF', 'GGF sem ACA', 'ACA sem MOB', 'MOB sem ACA', 'RRQ Sem REQ Item cor.', 'Restrição GGF/MOB', 'Últ Movto', 'Tipo']].to_html(index=False, classes="no-wrap"))
        f.write('</table>')
        f.write('</div>')
        f.write('<hr>')
        f.write('<br>')

    f.write(f'<small class="Col-md-12">Última data de movimentação no totvs: {data_mais_recente}</small>')

    # Escrever o gráfico de barras no arquivo HTML
    f.write('<div class="col-md-12 mt-3 justify-content-center">')
    f.write('<h3>Gráfico:</h3>')
    f.write(f'<img src="data:image/png;base64,{image_base64}"/>')
    f.write('</div>')

    f.write('</div>')

    f.write("</div>\n")
    f.write("</div>\n")

    # Incluir os scripts do Bootstrap no final do arquivo HTML
    f.write('<script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"></script>\n')
    f.write('<script src="https://cdn.jsdelivr.net/npm/@popperjs/core@2.5.4/dist/umd/popper.min.js"></script>\n')
    f.write('<script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>\n')
    f.write("</body>\n</html>")

print("Os resultados dos filtros foram salvos no arquivo HTML:", arquivoHTML)



