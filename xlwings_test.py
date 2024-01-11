from swxl.openpyxlpro.core import offset
from swxl.pandaspro.core import pwread
import pandas as pd
import xlwings as xw
from source import process_interim

template_cell = {
    'CA account deficit (US$)': 'G5',
    'CA account (in percent of GDP)':'G6',
    'Capital account (US$)': 'G7',
    'Financial accounts (US$)': 'G8',
    'Errors and omissions (US$)': 'G9',
    'Reserve accumulation (US$)': 'G10',
    'Any other below the line operations (US$)': 'G11',
    'Fiscal balance (in percent of GDP)': 'G13',
    'Exceptional financing (US$)': 'G15',
    'IMF': 'G16',
    'Other IFIs': 'G17',
    'Bilaterals': 'G18',
    'Others': 'G19'
}

bopterms = {
    'CA account deficit (US$)': ['current account balance', 'current account'],
    'CA account (in percent of GDP)':['CA account (in percent of GDP)', 'current account in percent of GDP', 'current account (in percent of GDP)'],
    'Capital account (US$)': ['capital account'],
    'Financial accounts (US$)': ['financial account'],
    'Errors and omissions (US$)': ['Errors and omissions'],
    'Reserve accumulation (US$)': ['Gross official reserves', 'Gross available reserves', 'Gross official foreign reserves', 'Gross international reserves'],
    'Any other below the line operations (US$)': [],
}
gfnterms = {
    'Fiscal balance (in percent of GDP)': [],
    'Exceptional financing (US$)': ['Exceptional financing', 'Financing need', 'Total financing needs', 'Official Financing'],
    'IMF': ['IMF'],
    'Other IFIs': [],
    'Bilaterals': [],
    'Others': ['other public sector']
}

currency_dict = {
    'CNY': 7
}

manual = r'C:\Users\xli7\International Monetary Fund (PRD)\SPR-Prolonged UFR - Documents\General\Data\ROC\Staff Reports Mining\external financing table file\final\config\manual_info.xlsx'
manual_data = pwread(manual)[0].set_index('index')

def getcell(wb, arrid, year, term, term_dict, manual_data):
    search_terms = term_dict[term]
    source = 'BOP' if term_dict == bopterms else 'GFN'
    find_sheet = f'{arrid}_{source}'
    try:
        ws_find = wb.sheets[find_sheet]
    except:
        return "not found"

    column, row = None, None
    find_year = ws_find.api.UsedRange.Find(year, LookAt=xw.constants.LookAt.xlWhole)
    if find_year is not None:
        column = find_year.Address.split('$')[1]
    for t in search_terms:
        found_cell = ws_find.api.UsedRange.Find(t, LookIn=-4163)
        if found_cell is not None:
            row = found_cell.Address.split('$')[2]
            break
    # If cell noted already in manual_info, use this
    try:
        cell = manual_data.at[term, str(arrid)] if not pd.isnull(manual_data.at[term, str(arrid)]) else "not found"
    except:
        raise KeyError(f"{arrid} not in the manual_info file, consider adding it")
    # Else check if found column + row both
    if cell == "not found" and column != None and row != None:
        # print(arrid, term, column, row)
        cell = column + str(row)
        # print(cell)

    return cell

def unit_bool(source, arrid):
    if manual_data.at[f'{source}_unit', str(arrid)] == 'm':
        return "millions"
    elif manual_data.at[f'{source}_unit', str(arrid)] == 'f':
        return "flag"
    else:
        return "billions"

def currency_exchange(source, arrid):
    if pd.isnull(manual_data.at[f'{source}_currency', str(arrid)]):
        return "Default USD"
    else:
        return manual_data.at[f'{source}_currency', str(arrid)]

def fillmaria(type='PRGT'):
    file = r'C:\Users\xli7\International Monetary Fund (PRD)\SPR-Prolonged UFR - Documents\General\Data\ROC\Staff Reports Mining\external financing table file\final' + f'/ROC2018_{type}.xlsx'
    wb = xw.Book(file)

    flags = {}
    count = 1
    data = process_interim(export=None).inlist('roc2018', type)
    for index, row in data.iterrows():
        arrid = row['program_num']
        year = row['approval_date'].year
        duration = round((row['actual_end_date'] - row['approval_date']).days / 365)

        print(f"{count}/{len(data)} - {arrid}")
        flags[arrid] = []

        try:
            ws = wb.sheets[str(arrid)]
        except:
            count += 1
            continue
        ws.range('B2').value = arrid
        ws.range('G3').value = year
        ws.range('F3').value = year - 1
        ws.range('H3').value = year + 1
        ws.range('I3').value = year + 2
        ws.range('J3').value = year + 3
        ws.range('M39').value = duration

        ## Flag 1: Unit
        if unit_bool('BOP', arrid) == 'flag' or unit_bool('GFN', arrid) == 'flag':
            unit_info = 'Unit flag: ' + f"BOP - {unit_bool('BOP', arrid)}, " +  f"GFN - {unit_bool('GFN', arrid)}"
            flags[arrid].append(unit_info)

        ## Flag 2: Currency
        if ((not pd.isnull(currency_exchange('BOP', arrid))) or (not pd.isnull(currency_exchange('GFN', arrid)))) \
                and currency_exchange('BOP', arrid) != currency_exchange('GFN', arrid):
            currency_info = 'Currency flag: ' + f"BOP - {currency_exchange('BOP', arrid)}, " +  f"GFN - {currency_exchange('GFN', arrid)}"
            flags[arrid].append(currency_info)

        ## Use loop to write down formula for each of the rows
        ################################
        def fillin(target_dict):
            source = 'BOP' if target_dict == bopterms else 'GFN'
            for t in target_dict.keys():
                cell = getcell(wb, arrid, year, t, target_dict, manual_data)
            ## Flag 3: SR row/column not found
                if cell == "not found":
                    cell_info = f"{t} not found in SR"
                    flags[arrid].append(cell_info)
                else:
                    # Decide about the adjustor
                    negative_adj = '-' if t in ['CA account deficit (US$)'] else ''
                    unit_adj = ''
                    currency_adj = ''

                    if unit_bool(source, arrid) == 'millions':
                        unit_adj = '/1000'
                    if currency_exchange(source, arrid) != 'Default USD':
                        try:
                            currency_adj = f'/{currency_dict[currency_exchange(source, arrid)]}'
                        except:
                            currency_adj = ''

                    if t in ['CA account (in percent of GDP)', 'Fiscal balance (in percent of GDP)']:
                        unit_adj = ''
                        currency_adj = ''

                    ws.range(template_cell[t]).value = f"={negative_adj}'{arrid}_{source}'!{cell}" + unit_adj + currency_adj
                    if t not in ['Exceptional financing (US$)', 'IMF', 'Other IFIs', 'Bilaterals', 'Others']:
                        ws.range(offset(template_cell[t], 0,-1)).value = f"={negative_adj}'{arrid}_{source}'!{offset(cell, 0, -1)}" + unit_adj + currency_adj
                    ws.range(offset(template_cell[t], 0, 1)).value = f"={negative_adj}'{arrid}_{source}'!{offset(cell, 0, 1)}" + unit_adj + currency_adj
                    ws.range(offset(template_cell[t], 0, 2)).value = f"={negative_adj}'{arrid}_{source}'!{offset(cell, 0, 2)}" + unit_adj + currency_adj
                    ws.range(offset(template_cell[t], 0, 3)).value = f"={negative_adj}'{arrid}_{source}'!{offset(cell, 0, 3)}" + unit_adj + currency_adj

        fillin(bopterms)
        fillin(gfnterms)
        count += 1
    return flags
    # ws = wb.sheets['724_BOP']
    # column = ws.api.UsedRange.Find(2018).Address.split('$')[1]
    # row = ws.api.UsedRange.Find('Current account balance', LookIn=-4163).Address.split('$')[2]


special_log = '''
682 GFN create 17
754 unmerge H and I column

'''

if __name__ == '__main__':
    a = fillmaria()


