# import tabula
# import tabulate
# import PyPDF2
from pdfminer.high_level import extract_text
from config import *
import re
# import os

gfn_key1 = ["Table"]
gfn_key2 = ["External Financing", "Gross Financing"]
gfn_key3 = ["Requirement", "Need"]
bop_key1 = ['Table']
bop_key2 = ['Balance of Payments', 'Balance Of Payments']
bop_key3 = ['Dollar', 'dollar']

def getpage_gfn(path_rawpdf):
    pattern = r'\d{3}_\d{3}'
    fileid = re.findall(pattern, path_rawpdf)[0]
    text = extract_text(path_rawpdf)
    matching_pages = []

    for page_number, page_text in enumerate(text.split('\x0C'), start=1):
        bool1 = all(keyword in page_text for keyword in gfn_key1)
        bool2 = any(keyword in page_text for keyword in gfn_key2)
        bool3 = any(keyword in page_text for keyword in gfn_key3)
        if bool1 and bool2 and bool3:
            matching_pages.append(page_number)

    matching_pages_gt_10 = [page for page in matching_pages if page > 10]
    if len(matching_pages_gt_10) == 0:
        matching_pages_gt_10 = [-999]
    elif len(matching_pages_gt_10) > 1:
        matching_pages_gt_10 = [correct_dict_gfn[fileid]]
    else:
        pass
    print(matching_pages_gt_10)
    return matching_pages_gt_10


def getpage_bop(path_rawpdf):
    pattern = r'\d{3}_\d{3}'
    fileid = re.findall(pattern, path_rawpdf)[0]
    text = extract_text(path_rawpdf)
    matching_pages = []

    for page_number, page_text in enumerate(text.split('\x0C'), start=1):
        bool1 = all(keyword in page_text for keyword in bop_key1)
        bool2 = any(keyword in page_text for keyword in bop_key2)
        bool3 = any(keyword in page_text for keyword in bop_key3)
        if bool1 and bool2 and bool3:
            matching_pages.append(page_number)

    matching_pages_gt_10 = [page for page in matching_pages if page > 10]
    if len(matching_pages_gt_10) == 0:
        matching_pages_gt_10 = [-999]
    elif len(matching_pages_gt_10) > 1:
        matching_pages_gt_10 = [correct_dict_bop[fileid]]
    else:
        pass
    print(matching_pages_gt_10)
    return matching_pages_gt_10