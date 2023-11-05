"""This module reads the projecten pages of the sia database for projects where
    the status = 'afgerond'. A total of some 127 project with 21 projects per
    page.
    For eacht project the project detail page is read and some 10 elements
    are extracted. The results are collected in a list. Ultimately the list
    is converted to a pandas dataframe and the dataframe is exported to excel.

Args:

Example:

Attributes:

To do:
     * add 'contact' field

"""

import pandas as pd
import requests
import os
from bs4 import BeautifulSoup

# initialize
row = []
df = pd.DataFrame([])


def write2excel(df: pd.DataFrame):
    # Source: https://xlsxwriter.readthedocs.io/
    # example_pandas_column_formats.html#ex-pandas-column-formats.

    with pd.ExcelWriter(os.path.join('tables', 'sia-projecten-afgerond.xlsx'),
                        engine='xlsxwriter') as writer:

        df.to_excel(writer,
                    sheet_name='SIA_database',
                    index=False)
        workbook = writer.book
        text_format = workbook.add_format({'text_wrap': True,
                                           'align': 'vcenter'})
        link_format = workbook.add_format({'color': 'blue',
                                           'underline': True,
                                           'text_wrap': True,
                                           'align': 'vcenter'})

        worksheet = writer.sheets['SIA_database']
        (max_row, max_col) = df.shape
        column_settings = [{"header": column} for column in df.columns]
        worksheet.add_table(0, 0, max_row, max_col - 1,
                            {"columns": column_settings})
        worksheet.set_column('A:A', 30, text_format)  # 1 Dossier
        worksheet.set_column('B:B', 30, text_format)  # 2 Titel
        worksheet.set_column('C:C', 10, text_format)  # 3 Status
        worksheet.set_column('D:D', 17, text_format)  # 4 Startdatum
        worksheet.set_column('E:E', 17, text_format)  # 5 Einddatum
        worksheet.set_column('F:F', 20, text_format)  # 6 Regeling
        worksheet.set_column('G:G', 20, link_format)  # 7 URL / Project
        worksheet.set_column('H:H', 99, text_format)  # 8 Lange beschrijving
        worksheet.set_column('I:I', 20, text_format)  # 9 Hogeschool
        worksheet.set_column('J:J', 50, text_format)  # 10 Thema ontbreekt soms
        worksheet.set_column('K:K', 20)               # Contactpersoon

# tabel programma?, tabel plaats?

        workbook.add_worksheet('Cover')
        worksheet = writer.sheets['Cover']
        worksheet.write(0, 0, 'Datum:')
        worksheet.write(1, 0, 'Input')
        worksheet.write(2, 0, 'Totaal projecten')
        worksheet.write(3, 0, 'Selectie')
        worksheet.write(4, 0, 'Geselecteerde projecten')  # Properties of df
        worksheet.write(5, 0, 'Output')
        worksheet.write(6, 0, 'Contact')


for page in range(0, 1):
    # read webpage(s), max page# = 127
    site = 'https://www.sia-projecten.nl'
    url = (site + '/zoek?key='
           '&status=Afgerond'
           '&programma='
           '&regeling='
           '&vhthemas='
           '&kennisinstelling='
           '&plaats='
           '&page='
           + str(page))

    reqs = requests.get(url)  # read projecten overview page
    soup = BeautifulSoup(reqs.text, 'lxml')

    projecten = soup.find_all('div', class_='view-project dm-teaser')
    for tag in projecten:
        titel = tag.find('h2').text
        project = site + tag.find('a', class_='meerlink').get('href')

        url = (project)  # read project detail page
        reqs = requests.get(url)
        soup = BeautifulSoup(reqs.text, 'lxml')
        row = []
        project_page = soup.find_all('tr')
        for tag in project_page:
            row.append(tag.findNext('td').text)

        row.insert(1, titel)
        row.insert(6, project)
        row.insert(7, soup.find('p', class_='samenvatting').text.strip())
        row.insert(8, soup.find('p', class_='hogeschool').text.strip())
        df = pd.concat([df, pd.DataFrame([row])])

df.columns = ['dossier', 'titel', 'status', 'begindatum',
              'einddatum', 'regeling', 'project', 'beschrijving',
              'hogeschool', 'themas']

write2excel(df)
