import shutil
from swxl.pandaspro.core import pwread
from create_pagedf import folder_extractpdf_root
from source import process_interim, path_final_roc2018_PRGT, path_final_roc2018_GRA, path_final_roc2024_PRGT, path_final_roc2024_GRA

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
def extract_excel():
    pass

# 3. consolidated 2 excels (using results from 2)
def consolidate():
    roc2018list = process_interim().inlist('ROC2018','GRA','PRGT')['srid']
    data = pwread(folder_extractpdf_root + '\index.xlsx').inlist('srid', roc2018list)

    file1 = ''
    for index, rows in data.iterrows():
        pass



if __name__ == '__main__':
    consolidate()