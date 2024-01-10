import pandas as pd
from swxl.pandaspro.core import pwread
from getpagenum import *
from extractpdf import *
from pdf2xl import *
import os
from source import process_interim

folder_rawpdf = r'C:\Users\xli7\International Monetary Fund (PRD)\SPR-Prolonged UFR - Documents\General\Data\ROC\Staff Reports Mining'
folder_extractpdf_root = r'C:\Users\xli7\International Monetary Fund (PRD)\SPR-Prolonged UFR - Documents\General\Data\ROC\Staff Reports Mining\external financing table file'
folder_savexl_root = r'C:\Users\xli7\International Monetary Fund (PRD)\SPR-Prolonged UFR - Documents\General\Data\ROC\Staff Reports Mining\external financing table file'
pagelist_gfn = {}
pagelist_bop = {}
folders = {'2018gra': '\ROC2018\GRA',
           '2018prgt': '\ROC2018\PRGT'}

def create_page_df(export=None):
    for index, row in process_interim().inlist('roc2018','GRA','PRGT').iterrows():
        file = row['path']
        # gfn dictionary
        pgnum_gfn = str(getpage_gfn(file)[0])
        pagelist_gfn[row['srid']] = pgnum_gfn

        # bop dictionary
        pgnum_bop = str(getpage_bop(file)[0])
        pagelist_bop[row['srid']] = pgnum_bop

    # Create the unioned new dataframe:
    # 1. this dataframe's id = srid + page number
    dfpagelist_gfn = pd.DataFrame({'srid': pagelist_gfn.keys(), 'pg': pagelist_gfn.values()})
    dfpagelist_bop = pd.DataFrame({'srid': pagelist_bop.keys(), 'pg': pagelist_bop.values()})
    d = pd.concat([dfpagelist_gfn, dfpagelist_bop], ignore_index=True, keys=['gfn', 'bop'])

    data = pwread(folder_rawpdf + r'\ROC_SR_source.xlsx')[0]
    accountmap = data.set_index('srid')['account'].to_dict()
    d['account'] = d.srid.map(accountmap)
    if export:
        d.to_excel(folder_extractpdf_root + r'\index.xlsx')
    return d

# if pdf:
#     path_extractpdf = os.path.join(folder_extractpdf, filename.replace('.pdf', ' ') + pgnum_gfn + '.pdf')
#     if pgnum_gfn != '-999':
#         export_1page(path_rawpdf, pgnum_gfn, folder_extractpdf)
#         time.sleep(2)
#         # pdf2xl(path_extractpdf, folder_savexl)

# if pdf:
#     path_extractpdf = os.path.join(folder_extractpdf, filename.replace('.pdf', ' ') + pgnum_bop + '.pdf')
#     if pgnum_bop != '-999':
#         export_1page(path_rawpdf, pgnum_bop, folder_extractpdf)
#         time.sleep(2)
#         pdf2xl(path_extractpdf, folder_savexl)

if __name__ == '__main__':
    d = create_page_df(export=True)