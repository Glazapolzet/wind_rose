import pandas as pd

import matplotlib.pyplot as plt

import numpy as np

from matplotlib.ticker import MultipleLocator, FormatStrFormatter, FixedLocator
from datetime import datetime
from windrose import WindroseAxes

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL


def df_preparation(b_w, snow=True, metel=True):
    print("def df_preparation")
    
    b_w['wd'] = b_w.wd.apply(lambda x: np.nan if x == 'CALM' else x)

    if snow:
        # выбор только тех значений, где sn=2
        b_w['sn'] = b_w.sn.apply(lambda x: np.nan if x != 2 else x)

    # b_w = b_w.dropna()
    # выбор только тех значений, где скорость ветра 3 или более
    if metel:
        b_w['ws'] = b_w.ws.apply(lambda x: np.nan if x < 3 else x)

    b_w = b_w.dropna()

    return b_w


def wind_rose(b_w, region, r_n, nw=False):
    # nw - признак того, что роза ветров строится по преимущественному направлению, независимо от скорости
    w_dir = {'N': 0, 'NNE': 22.5, 'NE': 45, 'ENE': 67.5, 'E': 90, 'ESE': 112.5,
             'SE': 135, 'SSE': 157.5, 'S': 180, 'SSW': 202.5, 'SW': 225,
             'WSW': 247.5, 'W': 270, 'WNW': 292.5, 'NW': 315, 'NNW': 337.5}

    b_w['wd'] = b_w.wd.apply(lambda x: w_dir[x])

    ax = WindroseAxes.from_ax()
    
    # nw=True значит только контур розы ветров
    # nw=false значит только или цветная или черно-белая роза ветров
    if nw:
        #только контур розы ветров
        b_w.ws = 4
        ax.contour(b_w.wd, b_w.ws, bins=np.arange(0, 16, 1), colors='k')
    else:
        if r_n == 0:
            #черно-белая роза ветров
            ax.contourf(b_w.wd, b_w.ws, bins=np.arange(0, 16, 1))
        elif r_n == 1:
            #цветная роза ветров
            ax.contour(b_w.wd, b_w.ws, bins=np.arange(0, 16, 1), colors='k')
        else:
            b_w.ws = 4
            ax.contour(b_w.wd, b_w.ws, bins=np.arange(0, 16, 1), colors='k')

    ax.set_xticklabels(['В', 'СВ', 'С', 'СЗ', 'З', 'ЮЗ', 'Ю', 'ЮВ'], fontsize=20, fontweight="bold")
    ax.set_theta_zero_location('E')
    #ax.set_legend()
    plt.savefig(region + '.jpg')
    
    return


def obr_file(archiv, date_n, date_k, snow, metel, r_n):
    print(archiv)
    print("obrfile")
    time_1 = [' 00:00', ' 03:00', ' 06:00', ' 09:00', ' 12:00', ' 15:00', ' 18:00', ' 21:00']
    print(archiv)
    file_csv = pd.read_csv(archiv, encoding='cp1251')

    print('134')
    dict = {
        'archiv_mcensk_05_24.csv': ['wrose_met Мценск за 10 лет.jpg',
                                    'Мценске', 'Мценск за период ', 'wrose Мценск'],
        'arсhiv_orel_05_24.csv': ['wrose_met Орел за 10 лет.jpg', 'Орле',
                                  'Орел за период ', 'wrose Орел'],
        'arсhiv_verh_05_21.csv': ['wrose_met Верховье за 10 лет.jpg', 'Верховье',
                                  'Верховье за период ',
                                  'wrose Верховье']}

    for i in range(8):
        date_1 = date_n + time_1[i]
        print(date_1)
        ind = file_csv.index[file_csv['dt_time'] == date_1].tolist()
        if len(ind) != 0:
            date_n = date_1
            break

    index_n = file_csv.loc[file_csv['dt_time'] == date_n].index
    for i in range(7, 0, -1):
        date_1 = date_k + time_1[i]
        ind = file_csv.index[file_csv['dt_time'] == date_1].tolist()
        if len(ind) != 0:
            date_k = date_1
            break
    index_k = file_csv.loc[file_csv['dt_time'] == date_k].index

    '''
если нет записи с временем 0:00, то нужно проверить на 03:00, затем на 06.00 
и так далее аналогично для date_k
    получение среза данных от начальной до конечной даты
    Поскольку файл отсортирован по убыванию даты и времени, ..............
    '''
    file_csv_daten_datek = file_csv.loc[index_k[0] - 1:index_n[0], :]
    print("filecsv")

    b_wind = pd.DataFrame()
    b_wind['wd'] = file_csv_daten_datek.Wind_dir
    b_wind['ws'] = file_csv_daten_datek.wind_speed
    b_wind['sn'] = file_csv_daten_datek.precipitation
    b_wind = df_preparation(b_wind, snow, metel)
    b_w = b_wind
    print(dict[archiv][1])
    region = dict[archiv][1] + ' ' + date_n + ' ' + date_k

    print('b_w',b_w)


    res = b_wind.groupby(['ws', 'wd']).size().reset_index(name='count')
    print(res['count'].sum())
    ss = res['count'].sum()
    print(b_wind.groupby(['ws','wd']).sum())

    wind_rose(b_wind, dict[archiv][3], r_n, False)
    #wind_rose(res, dict[archiv][3], r_n, False)
    deg = [ 0, 22.5, 45, 67.5, 90,112.5, 135, 157.5, 180, 202.5, 225, 247.5,
           270, 292.5, 315, 337.5]

    #func = lambda x: round(x.count())
    func = lambda x: round(100 * x.count() / res.shape[0], 2)
    #func = lambda x: round(100 * x.count() / ss, 2)
    pivot = pd.pivot_table(res, values='count', index=['ws'],margins=True,
                           columns=['wd'])

    p = pivot.fillna(0)
    print(p)
    #wind_rose(pivot, dict[archiv][3], r_n, False)
    if snow:
        sn = 'Осадки в виде снега, '
    else:
        sn = 'Независимо от осадков,'
    if metel:
        sn = sn + 'ветер 3 и более м/с. '
    else:
        sn = sn + 'независимо от скорости ветра.'
    region = dict[archiv][1] + '. Данные с ' + date_n + ' по ' \
             + date_k + '\n' + sn

    file_name = pivot_table_to_word(p, dict[archiv], region)
    return (file_name)


def pivot_table_to_word(pivot, l_list, region):
    print('pivot====================', pivot)
    #print(l_list)
    print('211')
    headers = ['С', 'ССВ', 'СВ', 'СВС', 'В', 'ВЮВ', 'ЮВ', 'ЮЮВ', 'Ю',
               'ЮЮЗ', 'ЮЗ', 'ЗЮЗ', 'З', 'ЗСЗ', 'СЗ', 'ССЗ', 'Всего,%']
    #deg = [0.0, 22.5, 45.0, 67.5,90.0, 112.5, 135.0, 157.5, 180.0,202.5, 225.0, 247.5, 270.0, 292.5, 315.0, 337.5,  'All']
    deg = ['N', 'NNE', 'NE', 'ENE', 'E', 'ESE','SE', 'SSE', 'S','SSW', 'SW','WSW', 'W', 'WNW', 'NW', 'NNW', 'All,%']
    document = Document()
    style = document.styles['Normal']
    font = style.font
    font.name = "Times New Roman"
    font.size = Pt(14)
    current_section = document.sections[-1]
    p = document.add_paragraph('Роза ветров в ')
    p.alignment = 1
    p.add_run(region)

    document.add_picture(l_list[3] + '.jpg', width=Inches(6))
    new_width, new_height = current_section.page_height, \
                            current_section.page_width
    new_section = document.add_section(WD_SECTION.NEW_PAGE)
    new_section.orientation = WD_ORIENT.LANDSCAPE
    new_section.page_width = new_width
    new_section.page_height = new_height

    style = document.styles['Normal']
    font = style.font
    font.name = "Times New Roman"
    font.size = Pt(10)



    table = document.add_table(rows=1, cols=18)
    table.style = 'Table Grid'

    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Напр. ветра/скорость ветра, м/с'

    li_0 = pivot.columns.tolist()
    #li_0=deg
    print('li0=',li_0)
    for j in range(1, 18):
        hdr_cells[j].text = headers[j - 1]
        hdr_cells[j].paragraphs[0].paragraph_format.alignment = \
            WD_TABLE_ALIGNMENT.CENTER
        hdr_cells[j].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    print(pivot.shape[0],pivot.shape[1])
    for i in range(0, pivot.shape[0]):

        row_cells = table.add_row().cells

        l = list(pivot.iloc[i])
        li = list(pivot.index)

        print('i=',i,'li=',li,l)

        li[len(li) - 1] = 'Всего,%'
        row_cells[0].text = str(li[i])
        # print(str(li[i]))
        row_cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
        # Ввиду того, что в сводной таблице pivot
        # могут отсутствовать некоторые направления ветра, вводится переменная j_sdvig
        #print('aadd')

        j_sdvig = 0
        for j in range(0, 16):

            if li_0[j - j_sdvig] == deg[j]:
                print('deg[j]=',deg[j])
                print('deg[j]=', deg[j], li_0[j - j_sdvig],j,j_sdvig)
                row_cells[j + 1].text = "{:.0f}".format(l[j - j_sdvig])
                row_cells[j + 1].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.RIGHT

                print("!!!!!!!!!!!!!!!!!!!!!!!!!!1")
                #print(row_cells[j + 1].text,)
            else:

                print("!!!!!!!!!!!!!!!!!!!!!!!!!!2")
                j_sdvig += 1
                print('deg[j]=', deg[j], li_0[j - j_sdvig])
                print('j=', j)
                #row_cells[j + 1].text = '0'

                row_cells[j + 1].text = "{:.0f}".format(l[j - j_sdvig])
                row_cells[j + 1].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.RIGHT

    print('282')

    try:
        document.save('Роза ветров в ' + l_list[1] +'. График и таблица.docx')

        file_name = 'Роза ветров в ' + l_list[1] + \
                    '. График и таблица.docx saved'

    except:
        document.save('Роза ветров в ' + l_list[1] + '. График и таблица1.docx')
        file_name = 'Роза ветров в ' + l_list[1] + '. График и таблица1.docx saved'
    return (file_name)