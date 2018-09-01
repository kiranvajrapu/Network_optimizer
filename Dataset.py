import pandas as pd
DL_file = ""
UL_file = ""
Sheet_First_part = ""
Sheet_Secound_part = ""
def get_file_DL(str):
    DL_file = str

def get_file_UL(str):
    UL_file = str

def get_file_UL(str):
    UL_file = str

def get_file_Sheet_First_part(str):
    Sheet_First_part = str

def get_file_Sheet_Secound_part(str):
    Sheet_Secound_part = str

def return_file_Sheet_First_part():
    return Sheet_First_part

def dataset_extract():
    print(Sheet_First_part)
    ds = pd.read_excel(Sheet_First_part)
    return ds
#
# return ds_Sheet_First_part
