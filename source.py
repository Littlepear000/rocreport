from swxl.pandaspro.core import pwread
import os
import pandas as pd
import shutil

path_report_ch = r'C:\Users\xli7\International Monetary Fund (PRD)\SPR-Prolonged UFR - Documents\General\Data\ROC\Staff Reports Mining\source files management\allreports_fromCH'
path_report_web = r'C:\Users\xli7\International Monetary Fund (PRD)\SPR-Prolonged UFR - Documents\General\Data\ROC\Staff Reports Mining\source files management\morereports_fromWEB'
path_final_roc2018_GRA = r'C:\Users\xli7\International Monetary Fund (PRD)\SPR-Prolonged UFR - Documents\General\Data\ROC\Staff Reports Mining\ROC2018\GRA'
path_final_roc2018_PRGT = r'C:\Users\xli7\International Monetary Fund (PRD)\SPR-Prolonged UFR - Documents\General\Data\ROC\Staff Reports Mining\ROC2018\PRGT'
path_final_roc2024_GRA = r'C:\Users\xli7\International Monetary Fund (PRD)\SPR-Prolonged UFR - Documents\General\Data\ROC\Staff Reports Mining\ROC2024\GRA'
path_final_roc2024_PRGT = r'C:\Users\xli7\International Monetary Fund (PRD)\SPR-Prolonged UFR - Documents\General\Data\ROC\Staff Reports Mining\ROC2024\PRGT'


def chget():
    srch_path = path_report_ch
    srch_names = []

    for file in os.listdir(srch_path):
        if file.endswith('.pdf'):
            srch_names.append(file)

    dfsrch = pd.DataFrame(srch_names, columns=['srid'])
    dfr0 = dfsrch[dfsrch['srid'].str.contains('R0')]
    dfr0['srid'] = dfr0['srid'].str.replace('_R0.pdf', '')
    return dfr0

def interim(export=None):
    rocxlsx = r'C:\Users\xli7\International Monetary Fund (PRD)\SPR-Prolonged UFR - Documents\General\Data\ROC\ROC base sample.xlsx'
    roc2018, mapping = pwread(rocxlsx, sheet_name='ROC2018')
    roc2024, mapping = pwread(rocxlsx, sheet_name='ROC2024')
    roc2018col = roc2018.columns.to_list()
    roc2024 = roc2024[roc2018col]

    combine = roc2018.merge(roc2024, on=roc2018col, how='outer', indicator=True, suffixes=('', '_roc2024'))
    combine.loc[(combine.inlist('_merge', ['left_only', 'both'], otype='m')) & (combine['account'] == 'GRA'), 'roc2018'] = 'GRA'
    combine.loc[(combine.inlist('_merge', ['left_only', 'both'], otype='m')) & (combine['account'] == 'PRGT'), 'roc2018'] = 'PRGT'
    combine.loc[(combine.inlist('_merge', ['right_only', 'both'], otype='m')) & (combine['account'] == 'GRA'), 'roc2024'] = 'GRA'
    combine.loc[(combine.inlist('_merge', ['right_only', 'both'], otype='m')) & (combine['account'] == 'PRGT'), 'roc2024'] = 'PRGT'
    combine.drop('_merge', axis=1, inplace=True)
    combine['srid'] = combine['country_code'].astype(str) + '_' +  combine['arrangement_number'].astype(str)
    chs = chget()
    combine = combine.merge(chs, on='srid', how='left', indicator=True)
    combine.loc[combine['_merge'] == 'both', 'source'] = 'CH'
    combine.drop('_merge', axis=1, inplace=True)
    combine['websearch'] = combine['country_name'] + ' ' + combine['approval_date'].dt.year.astype(str) + ' request ' + \
                           combine['arrangement_type']
    websearch = combine.loc[combine['source'].isna(), ['websearch', 'srid']]
    if export:
        websearch.to_excel('websearch.xlsx')
    return combine
# After success full web seach, interim will include all pdfs

def process_interim(export=None):
    combine = interim()
    srweb_path = path_report_web
    srweb_names = []

    for file in os.listdir(srweb_path):
        if file.endswith('.pdf'):
            srweb_names.append(file)

    dfsrweb = pd.DataFrame(srweb_names, columns=['srid'])
    dfsrweb['srid'] = dfsrweb['srid'].str.replace('.pdf', '')

    combine = combine.merge(dfsrweb, on='srid', how='left', indicator=True)
    combine.loc[combine['_merge'] == 'both', 'source'] = 'WEB'
    combine.drop(['_merge'], axis=1, inplace=True)
    combine = combine.sort_values(by='approval_date')
    combine.rename(columns={'arrangement_number': 'program_num', 'arrangement_type': 'facility'}, inplace=True)
    combine['path'] = combine.apply(lambda row: path_report_ch + f"/{row['srid']}_R0.pdf" if row[
                                                                                                 'source'] == 'CH' else path_report_web + f"/{row['srid']}.pdf",
                                    axis=1)
    combine['num'] = combine.reset_index().index + 2
    combine['hyper'] = combine['num'].apply(lambda cell: f'=HYPERLINK(L{cell},J{cell}& ".pdf")')
    combine.drop('num', axis=1, inplace=True)
    if export:
        combine.export_excel('ROC_SR_source.xlsx')
    return combine

if __name__ == '__main__':
    interim(export=True)