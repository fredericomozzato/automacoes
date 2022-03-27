import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from datetime import datetime
import csv

# definir ano e mes atual
currentMonth = datetime.now().month
currentYear = datetime.now().year

# carregar dados do YT .csv
csvFile = open('Dados da tabela.csv')
csvReader = csv.DictReader(csvFile)
ytData = list(csvReader)
# csvWriter =

videoList = []

wd_data = openpyxl.load_workbook('Formulário WD-40.xlsx')
wd_sheet = wd_data['Abril']

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


wd_data.save('/Users/fredericomozzato/Desktop/FormulárioWD_CSV.xlsx')
