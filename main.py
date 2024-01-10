import time

import pandas as pd

from config import *
from getpagenum import *
from extractpdf import *
from pdf2xl import *
import os
import re

folder_rawpdf = r'C:\Users\xli7\International Monetary Fund (PRD)\SPR-Prolonged UFR - Documents\General\Data\ROC\Staff Reports Mining'
folders = ['\ROC2018\PRGT']
folder_extractpdf_root = r'C:\Users\xli7\International Monetary Fund (PRD)\SPR-Prolonged UFR - Documents\General\Data\ROC\Staff Reports Mining\external financing table file'
folder_savexl_root = r'C:\Users\xli7\International Monetary Fund (PRD)\SPR-Prolonged UFR - Documents\General\Data\ROC\Staff Reports Mining\external financing table file'
pagelist_gfn = {}
pagelist_bop = {}

for folder in folders:
    path = folder_rawpdf + folder
    folder_extractpdf = folder_extractpdf_root + folder
    folder_savexl = folder_savexl_root + folder
    for filename in os.listdir(path):
        if filename.endswith(".pdf"):
            print(filename)
            path_rawpdf = os.path.join(path, filename)
            pgnum_gfn = str(getpage_gfn(path_rawpdf)[0])
            pagelist_gfn[filename.replace('.pdf', '')] = pgnum_gfn
            path_extractpdf = os.path.join(folder_extractpdf, filename.replace('.pdf', ' ') + pgnum_gfn + '.pdf')
            if pgnum_gfn != '-999':
                export_1page(path_rawpdf, pgnum_gfn, folder_extractpdf)
                time.sleep(2)
                pdf2xl(path_extractpdf, folder_savexl)
            else:
                pass

            pgnum_bop = str(getpage_bop(path_rawpdf)[0])
            pagelist_bop[filename.replace('.pdf', '')] = pgnum_bop
            path_extractpdf = os.path.join(folder_extractpdf, filename.replace('.pdf', ' ') + pgnum_bop + '.pdf')
            if pgnum_bop != '-999':
                export_1page(path_rawpdf, pgnum_bop, folder_extractpdf)
                time.sleep(2)
                pdf2xl(path_extractpdf, folder_savexl)
            else:
                pass
dfpagelist_gfn = pd.DataFrame(pagelist_gfn)
dfpagelist_gfn.to_excel(folder_extractpdf_root + r'\ROC2018\pagelist_gfn.xlsx')
dfpagelist_bop = pd.DataFrame(pagelist_bop)
dfpagelist_bop.to_excel(folder_extractpdf_root + r'\ROC2018\pagelist_bop.xlsx')