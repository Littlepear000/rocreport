import shutil
import os
import time
from swxl.pandaspro.core import pwread

import xlwings as xw

from create_pagedf import folder_external_financing, folder_rawpdf
from extractpdf import export_1page
from movetabs import copy_move_tab
from pdf2xl import pdf2xl
from source import process_interim, path_final_roc2018_PRGT, path_final_roc2018_GRA, path_final_roc2024_PRGT, path_final_roc2024_GRA

template = r'C:\Users\xli7\International Monetary Fund (PRD)\SPR-Prolonged UFR - Documents\General\Data\ROC\Staff Reports Mining\external financing table file\final\config\template_sheet.xlsx'
gra18_init = r'C:\Users\xli7\International Monetary Fund (PRD)\SPR-Prolonged UFR - Documents\General\Data\ROC\Staff Reports Mining\external financing table file\ROC2018_GRA.xlsx'
prgt18_init = r'C:\Users\xli7\International Monetary Fund (PRD)\SPR-Prolonged UFR - Documents\General\Data\ROC\Staff Reports Mining\external financing table file\ROC2018_PRGT.xlsx'
gra18_final = r'C:\Users\xli7\International Monetary Fund (PRD)\SPR-Prolonged UFR - Documents\General\Data\ROC\Staff Reports Mining\external financing table file\final\ROC2018_GRA.xlsx'
prgt18_final = r'C:\Users\xli7\International Monetary Fund (PRD)\SPR-Prolonged UFR - Documents\General\Data\ROC\Staff Reports Mining\external financing table file\final\ROC2018_PRGT.xlsx'

# 1. moved folder
def move():
    combine = process_interim()
    for index, row in combine.iterrows():
        if row['roc2018'] == "PRGT":
            print(f"Copying {row['srid']}.pdf")
            shutil.copy(row['path'], path_final_roc2018_PRGT + f"/{row['srid']}.pdf")
        elif row['roc2018'] == "GRA":
            print(f"Copying {row['srid']}.pdf")
            shutil.copy(row['path'], path_final_roc2018_GRA + f"/{row['srid']}.pdf")
        else:
            pass

        if row['roc2024'] == "PRGT":
            print(f"Copying {row['srid']}.pdf")
            shutil.copy(row['path'], path_final_roc2024_PRGT + f"/{row['srid']}.pdf")
        elif row['roc2024'] == "GRA":
            print(f"Copying {row['srid']}.pdf")
            shutil.copy(row['path'], path_final_roc2024_GRA + f"/{row['srid']}.pdf")
        else:
            pass

# 2. extracted excels (with 1-page PDF)
def extract_excel(path_rawpdf, pgnum_gfn, folder_extractpdf, outid):
    export_1page(path_rawpdf, pgnum_gfn, folder_extractpdf)
    time.sleep(2)
    pdf2xl(outid, folder_extractpdf)

# 3. consolidated 2 excels (using results from 2)
# Remember to create tables first
def final_indexpage_create():
    # initiate
    shutil.copy(prgt18_init, prgt18_final)
    shutil.copy(gra18_init, gra18_final)

    t1 = folder_external_financing + f"/final/ROC2018_GRA.xlsx"
    t2 = folder_external_financing + f"/final/ROC2018_PRGT.xlsx"

    source = folder_rawpdf + r'\ROC_SR_source.xlsx'
    index = folder_external_financing + r'\index.xlsx'
    keeplist = ['program_num', 'country_name', 'country_code', 'facility', 'approval_date', 'actual_end_date',
                'account', 'srid', 'GFN table', 'BOP table']

    sourcedata = pwread(source)[0]
    data = pwread(index)[0]
    data['table'] = data.apply(lambda row: "=HYPERLINK(\"#'" + row['srid'][-3:] + '_' + row['source'] + f"'!A1\", \"{row['srid'][-3:] + '_' + row['source']}\")" if row['pg'] != -999 else row['srid'][-3:] + '_' + row['source'], axis=1)
    data.drop('pg', axis=1, inplace=True)
    data = data.pivot(index=['srid'], columns='source', values='table').reset_index()
    data = data.merge(sourcedata, on='srid', how='left')
    data.rename(columns={'BOP': 'BOP table', 'GFN': 'GFN table'}, inplace=True)
    df_prgt = data.inlist('account', 'PRGT')[keeplist]
    df_gra = data.inlist('account', 'GRA')[keeplist]
    df_gra.export_excel(t1, '1_index')
    df_prgt.export_excel(t2, '1_index')

def consolidate(addmore=False):
    roc2018list = process_interim().inlist('roc2018','GRA','PRGT')['srid'].to_list()
    data = pwread(folder_external_financing + '\index.xlsx')[0].inlist('srid', roc2018list).inlist('pg', -999, invert=True)
    data['sort'] = data['srid'].str[-3:]
    data = data.sort_values(['sort'])
    i = 1
    pre_sort = None
    for index, row in data[::-1].iterrows():
        print(f"{i}/{len(data)}")
        curr_sort = row['sort']
        sourcefile = f"{folder_external_financing}/ROC2018/{row['account']}/{row['srid']} {row['pg']}.xlsx"
        if os.path.exists(sourcefile):
            if addmore==False:
                tabname = f"{row['srid'][-3:]}_{row['source']}"
                templatename = f"{row['srid'][-3:]}"
                targetfile = folder_external_financing + f"/final/ROC2018_{row['account']}.xlsx"
                print(row['srid'])
                print('=============================')
                tb = xw.Book(targetfile)
                if pre_sort is None or curr_sort != pre_sort:
                    copy_move_tab(template, 'template', templatename, targetwb=tb)
                    time.sleep(1)
                copy_move_tab(sourcefile, 'Table 1', tabname, targetwb=tb)
                print('tab move completed!')
                pre_sort = row['sort']
            else:
                pass
        else:
            print(f"{row['srid']}'s {row['account']} BOP is on page {row['pg']}")
            path = folder_rawpdf + f"/ROC2018/{row['account']}/{row['srid']}.pdf"
            folder_extractpdf = folder_external_financing + "\ROC2018" + "\\" + row['account']
            fileid = row['srid'] + ' ' + str(row['pg'])
            extract_excel(path, str(row['pg']), folder_extractpdf, fileid)
        i+=1


if __name__ == '__main__':
    final_indexpage_create()
    consolidate()
