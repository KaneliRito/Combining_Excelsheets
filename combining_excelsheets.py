#! /usr/bin/python
import time
import xlwings as xw
from pathlib import Path

def Combine_worksheet():
    BASE_DIR = Path(__file__).parent
    SOURCE_DIR = BASE_DIR
    # Makes a new folder called output where the combined excelsheet will be stored.
    OUTPUT_DIR = BASE_DIR / 'output'

    # Create output directory
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # Selects all excel files in current folder.
    excel_files = Path(SOURCE_DIR).glob('*.xlsx')

    # Create timestamp
    t = time.localtime()
    # Timeformat
    timestamp = time.strftime('%Y-%m-%d_%H;%M;%S', t)


    with xw.App(visible=False) as app:
        combined_wb = app.books.add()
        for excel_file in excel_files:
            wb = app.books.open(excel_file)
            for sheet in wb.sheets:
                sheet.copy(after=combined_wb.sheets[0])
            wb.close()
        combined_wb.sheets[0].delete()
        #name your worksheet
        #timestamp included in this version
        combined_wb.save(OUTPUT_DIR / f'all_worksheets_{timestamp}.xlsx')
        combined_wb.close()

Combine_worksheet()