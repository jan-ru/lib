""" This module reads the projecten pages of the sia database for projects where
    the status = 'afgerond'. A total of some 127 pages with 21 projects per
    page.
    For eacht project the project detail page is read and some 10 elements
    are extracted. The results are collected in a list. Ultimately the list
    is converted to a pandas dataframe and the dataframe is exported to excel.

Args:
    website (str): currently only 'https://www.sia-projecten.nl'
    excel_file (str): currently only 'sia-projecten-afgerond.xlsx'

Example:
    research_projects('website', 'excel_file')

Attributes:

To do:
     add 'contact' field
     www.nro.nl (scraping?)
     www.nwo.nl (api)
     add dataset metadata (docs.profiling.ydata.ai)

"""

import pandas as pd
import requests
import os
import datetime
from bs4 import BeautifulSoup
from typing import List

# initialize
row: List[int] = []
df = pd.DataFrame([])
input_website = "https://www.sia-projecten.nl"
output_file = "sia-projecten-afgerond.xlsx"
contact_email = "j.r.muller@hva.nl"


def write2excel_table(
    df: pd.DataFrame, input_website: str, output_file: str, sheet: str
):
    """
    Source: https://xlsxwriter.readthedocs.io/
    example_pandas_column_formats.html#ex-pandas-column-formats.
    """

    with pd.ExcelWriter(
        os.path.join("tables", output_file), engine="xlsxwriter"
    ) as writer:
        df.to_excel(writer, sheet_name=sheet, index=False)
        workbook = writer.book
        # print('type workbook: ', workbook)
        text_format = workbook.add_format({"text_wrap": True, "align": "vcenter"})
        link_format = workbook.add_format(
            {"color": "blue", "underline": True, "text_wrap": True, "align": "vcenter"}
        )

        worksheet = writer.sheets[sheet]
        (max_row, max_col) = df.shape
        column_settings = [{"header": column} for column in df.columns]
        worksheet.add_table(0, 0, max_row, max_col - 1, {"columns": column_settings})
        worksheet.set_column("A:A", 30, text_format)  # 1 Dossier
        worksheet.set_column("B:B", 30, text_format)  # 2 Titel
        worksheet.set_column("C:C", 10, text_format)  # 3 Status
        worksheet.set_column("D:D", 17, text_format)  # 4 Startdatum
        worksheet.set_column("E:E", 17, text_format)  # 5 Einddatum
        worksheet.set_column("F:F", 20, text_format)  # 6 Regeling
        worksheet.set_column("G:G", 20, link_format)  # 7 URL / Project
        worksheet.set_column("H:H", 99, text_format)  # 8 Lange beschrijving
        worksheet.set_column("I:I", 20, text_format)  # 9 Hogeschool
        worksheet.set_column("J:J", 50, text_format)  # 10 Thema ontbreekt soms
        worksheet.set_column("K:K", 20)  # Contactpersoon

        #    write2excel_cover(writer, workbook, input_website, output_file,
        #    'Cover', 'j.r.muller@hva.nl', max_row, max_col)

        # tabel programma?, tabel plaats?

        # writer
        # def write2excel_cover(writer: pd.ExcelWriter,
        # workbook: pd.ExcelWriter.book,
        # input_website: str, output_file: str, sheet: str, contact_email: str,
        # max_row: int, max_col: int):
        # workbook = writer.book
        workbook.add_worksheet("Cover")
        worksheet = writer.sheets["Cover"]
        worksheet.set_column("A:A", 25)  # Labels
        worksheet.set_column("B:B", 30)  # Labels
        worksheet.write_row(
            0, 0, ("Datum:", datetime.date.today().strftime("%d %B %Y"))
        )
        worksheet.write_row(1, 0, ("Input:", input_website))
        worksheet.write_row(2, 0, ("Output:", output_file))
        worksheet.write_row(3, 0, ("Selectie:", "status=afgerond"))
        worksheet.write_row(4, 0, ("Dataframe kolommen:", str(max_col)))
        worksheet.write_row(5, 0, ("Dataframe regels:", str(max_row)))
        worksheet.write_row(6, 0, ("", ""))
        worksheet.write_row(7, 0, ("Contact", contact_email))


# eventueel unique values per kolom?

for page in range(0, 1):
    # read webpage(s), max page# = 127
    url = (
        input_website + "/zoek?key="
        "&status=Afgerond"
        "&programma="
        "&regeling="
        "&vhthemas="
        "&kennisinstelling="
        "&plaats="
        "&page=" + str(page)
    )

    reqs = requests.get(url)  # read projecten overview page
    soup = BeautifulSoup(reqs.text, "lxml")

    projecten = soup.find_all("div", class_="view-project dm-teaser")
    for tag in projecten:
        titel = tag.find("h2").text
        project = input_website + tag.find("a", class_="meerlink").get("href")

        url = project  # read project detail page
        reqs = requests.get(url)
        soup = BeautifulSoup(reqs.text, "lxml")
        row = []
        project_page = soup.find_all("tr")
        for tag in project_page:
            row.append(tag.findNext("td").text)

        row.insert(1, titel)
        row.insert(6, project)
        row.insert(7, soup.find("p", class_="samenvatting").text.strip())
        row.insert(8, soup.find("p", class_="hogeschool").text.strip())
        df = pd.concat([df, pd.DataFrame([row])])

df.columns = [
    "dossier",
    "titel",
    "status",
    "begindatum",
    "einddatum",
    "regeling",
    "project",
    "beschrijving",
    "hogeschool",
    "themas",
]

write2excel_table(df, input_website, output_file, "SIA-database")
# write2excel_cover(input_webiste, output_file, 'Cover', contact_email)
# write2excel_contact_details(from file or from header)
# write2excel_df_describe(df,writer,sheet)
