from swxl.pandaspro.core import pwread
import os
import pandas as pd
import roc_pdf2xl.config as c

def chget():
    srch_path = c.path_report_ch
    srch_names = []

    for file in os.listdir(srch_path):
        if file.endswith('.pdf'):
            if 'R0' in file:
                srch_names.append(file.replace('_R0.pdf', ''))

    dfsrch = pd.DataFrame(srch_names, columns=['srid'])
    return dfsrch

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
    srweb_path = c.path_report_web
    srweb_names = []

    for file in os.listdir(srweb_path):
        if file.endswith('.pdf'):
            srweb_names.append(file.replace('.pdf', ''))

    dfsrweb = pd.DataFrame(srweb_names, columns=['srid'])

    combine = combine.merge(dfsrweb, on='srid', how='left', indicator=True)
    combine.loc[combine['_merge'] == 'both', 'source'] = 'WEB'
    combine.drop(['_merge'], axis=1, inplace=True)
    combine = combine.sort_values(by='approval_date')
    combine.rename(columns={'arrangement_number': 'program_num', 'arrangement_type': 'facility'}, inplace=True)
    combine['path'] = combine.apply(lambda row: c.path_report_ch + f"/{row['srid']}_R0.pdf" if row['source'] == 'CH' else c.path_report_web + f"/{row['srid']}.pdf",
                                    axis=1)
    combine['num'] = combine.reset_index().index + 2
    combine['hyper'] = combine['num'].apply(lambda cell: f'=HYPERLINK(L{cell},J{cell}& ".pdf")')
    combine.drop('num', axis=1, inplace=True)
    if export:
        combine.export_excel('ROC_SR_source.xlsx')
    return combine


if __name__ == '__main__':
    # interim(export=True)
    data = process_interim()
    data.drop('websearch', axis=1, inplace=True)
    index = pwread('./roc_pdf2xl/meta_raw/index.xlsx')[0]
    gfn = index.inlist('source','GFN').set_index('srid')['pg'].to_dict()
    bop = index.inlist('source', 'BOP').set_index('srid')['pg'].to_dict()
    data['gfn'] = data['srid'].map(gfn)
    data['bop'] = data['srid'].map(bop)
    data.export_excel('./roc_pdf2xl/meta.xlsx', 'dashboard','A1', structure='sharelist1')


