'''
Os imports principais.
'''
import datetime as dt
import sys
from os import getcwd, chdir
from os.path import join
from openpyxl import load_workbook
from winotify import Notification

atual_dir = getcwd()
sys.path.append(atual_dir)
if getattr(sys, "frozen", False):
    chdir(sys._MEIPASS)

arquivo_excel = join(atual_dir, 'example.xlsx')
icon = join(atual_dir, 'icon.png')

wb = load_workbook(arquivo_excel)
sheet = wb.active

agora = (dt.datetime.now()).date()
eventos_proximos = []

for row in sheet.iter_rows(min_row=2):
    data = (row[0].value).date()
    evento = row[1].value
    tipo = row[2].value

    # Ver qual é a data comemorativa mais próxima
    if agora <= data <= agora + dt.timedelta(days=7) or data == agora:
        eventos_proximos.append((evento, data.strftime("%d/%m/%Y"), tipo))

# Adicionar um verificador que todo o dia notifica caso uma data esteja próxima
if eventos_proximos:
    EVENTO = ', '.join(f'{evento} em {data} do tipo {tipo}' for evento, data, tipo in eventos_proximos)
    notif = Notification(app_id='Notificador de Eventos',
                         title='Datas Comemorativas',
                         msg=f'Olá, os eventos mais próximos são: {EVENTO}',
                         duration='long',
                         icon=icon
                         )
    notif.show()
# 7 dias antes, as 11:00 e as 16:00
