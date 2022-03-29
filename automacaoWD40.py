'''
Automacao para preenchimento de relatorios WD-40
Autor: Frederico Mozzato, mar/2022
'''

import openpyxl
from datetime import datetime
import csv
import sys
import os
import subprocess

# definir arquivos via linha de comando
tabelaYt = sys.argv[1]
mes = sys.argv[2]

# referencia para meses:
monthRefer = {
    '1': 'Janeiro',
    '2': 'Fevereiro',
    '3': 'Março',
    '4': 'Abril',
    '5': 'Maio',
    '6': 'Junho',
    '7': 'Julho',
    '8': 'Agosto',
    '9': 'Setembro',
    '10': 'Outubro',
    '11': 'Novembro',
    '12': 'Dezembro'
}

# definir ano e mes atual
currentMonth = datetime.now().month
currentYear = datetime.now().year

# carregar dados do YT .csv
csvFile = open(tabelaYt)
csvReader = csv.DictReader(csvFile)
ytData = list(csvReader)
# csvWriter =

videoList = []

wd_data = openpyxl.load_workbook('Formulário WD-40.xlsx')
wd_sheet = wd_data[monthRefer[mes]]

for row in ytData[1:]:
    if row['Horário de publicação do vídeo'] is None:
        continue
    row['Horário de publicação do vídeo'] = datetime.strptime(row['Horário de publicação do vídeo'], '%b %d, %Y')

for row in ytData[1:]:
    if row['Horário de publicação do vídeo'] is None:
        continue
    if row['Horário de publicação do vídeo'].month == currentMonth and row['Horário de publicação do vídeo'].year == currentYear:
        videoList.append(row)


videoListSorted = sorted(videoList, key=lambda x: x['Horário de publicação do vídeo'], reverse=True)

# loop para inserir valores na tabela
counter = 0
for video in videoListSorted:
    wd_sheet['A' + str(4 + counter)] = 'Youtube'
    wd_sheet['B' + str(4 + counter)] = '@foradaspistas'
    wd_sheet['C' + str(4 + counter)] = 'Vídeo'
    wd_sheet['D' + str(4 + counter)] = 'EVD'
    wd_sheet['E' + str(4 + counter)].hyperlink = 'https://youtube.com/watch?v=' + video['Vídeo']
    wd_sheet['F' + str(4 + counter)] = video['Horário de publicação do vídeo'].strftime('%d/%m/%y')
    wd_sheet['G' + str(4 + counter)] = int(video['Impressões'])
    wd_sheet['H' + str(4 + counter)] = None
    wd_sheet['I' + str(4 + counter)] = int(video['Marcações "Gostei"'])
    counter += 1

# salvar arquivo final na pasta do programa
wd_data.save(f'Formulário WD-40_{monthRefer[mes]}.xlsx')

# referencia de arquivo para abrir a tabela
outputFile = f'Formulário WD-40_{monthRefer[mes]}.xlsx'

# logica para abrir o arquivo em qualquer sistema operacional
if sys.platform == 'win32':
    os.startfile(outputFile)
else:
    opener = "open" if sys.platform == 'darwin' else 'xdg-open'
    subprocess.call([opener, outputFile])
