import tabula
import pandas
import os
import numpy as np
import pandas as pd
import xlwings as xw
import logging
import time
from datetime import date
import xlwings.constants as win32c

def num_to_col_letters(num):
    try:
        letters = ''
        while num:
            mod = (num - 1) % 26
            letters += chr(mod + 65)
            num = (num - 1) // 26
        return ''.join(reversed(letters))
    except Exception as e:
        raise e
output_location= os.getcwd()
today_date=date.today()
download_path=os.getcwd() + "\\Download"
file_name= os.listdir(os.getcwd() + "\\Download")
file2=file_name[0]
test_area_date = ["67,47,85,160"]
df_date = tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages="1",area=test_area_date,silent=True,guess=False)
Trade_date=df_date[0].columns[1]

def trim_process_csv(structure_name,file_name:str):
    try:
        input_sheet=os.getcwd()+f'\\{file_name}' 
        # logger.info("Opening operating workbook instance of excel")
        retry=0
        while retry < 10:
            try:
                wb=xw.Book(input_sheet)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e
        ws1=wb.sheets[0]
        ws1.autofit()
        ws1.api.Range("A1").EntireColumn.Delete()
        ws1.range("1:1").api.Delete()
        list1=["CONTRACT","OPTION_EXPIRY DATE","TRADE_DAYS_TO_EXPIRY","FUTURES_PRICE","ATM_STRADDLE","BREAK_EVEN","ATM_IMP_VOL","ATM_IMP_VOL_1D_CHG","ATM_IMP_VOL_1WK_CHG","ATM_IMP_VOL_1MO_CHG","REAL_VOL_AVG_HIST","REAL_VOL_30_DAY","IMP_VOL_10D_PUT","IMP_VOL_25D_PUT","IMP_VOL_ATM","IMP_VOL_25D_CALL","IMP_VOL_10D_CALL"]
        column_list = ws1.range("A1").expand('right').value
        for index2,values in enumerate(column_list):
            values_column_no=column_list.index(values)+1
            values_letter_column = num_to_col_letters(column_list.index(values)+1)
            ws1.range(f"{values_letter_column}1").value = list1[index2]                  
        # print("Pause")
        ws1.autofit()
        last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        ws1.api.Range("A1").EntireColumn.Insert()
        ws1.range("A1").value="STRUCTURE_NAME"
        ws1.range(f"A2:A{last_row}").value=structure_name
        ws1.api.Range("A1").EntireColumn.Insert()
        ws1.range("A1").value="TRADE_DATE"
        ws1.range(f"A2:A{last_row}").value=Trade_date
        column_list = ws1.range("A1").expand('right').value
        insert_letter=num_to_col_letters(len(column_list)+1)
        ws1.range(f"{insert_letter}1").value="INSERTDATE"
        ws1.range(f"{insert_letter}2:{insert_letter}{last_row}").value=today_date
        insert_letter=num_to_col_letters(len(column_list)+2)
        ws1.range(f"{insert_letter}1").value="UPDATEDATE"
        ws1.api.Range("S:S").Cut()
        ws1.api.Range("R:R").Insert(win32c.Direction.xlToRight)
        wb.save(f"{output_location}\\{file_name}")
        try:
            wb.app.quit()
        except Exception as e:
            wb.app.quit()          
    except Exception as e:
                    pass  
       
def read_pdf():
    try:
        column_values0=["82","114","151","181","213","242","273","296","327","357","391","420","450","481","509","539"]
        test_area0 = ["293.378,50,526.173,569.543"]
        column_values1=["82","114","151","181","213","242","272","296","327","357","391","420","450","481","509","539"]
        test_area1=["282.996","50.0","506.376","567.905"]
        test_area2=["530.0","50.0","584","568"]
        column_values2=["82","114","151","181","213","242","272","296","327","357","391","420","450","481","509","539"]
        df = tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages="all",silent=True,guess=False)
        for index, page in enumerate(df):
            if index == 0:
                test_area_structure = ["291,48,304,569"]
                df_structure = tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages="1",area=test_area_structure,silent=True,guess=False)
                structure_name=df_structure[0].columns[0]
                df0=tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages=[index + 1],area=test_area0,columns=column_values0,silent=True,guess=False)    
                file_name='file.csv'
                df0[0].to_csv(file_name)
                trim_process_csv(structure_name,file_name)
                #for second table
                test_area_structure2 = ["524.408,37.026,542,588.591"]
                df_structure2 = tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages="1",area=test_area_structure2,silent=True,guess=False)
                structure_name=df_structure2[0].columns[0]
                df1=tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages=[index + 1],area=test_area2,columns=column_values2,silent=True,guess=False)    
                file_name='file1.csv'
                df1[0].to_csv(file_name)
                # df1[0].drop(0,inplace=True)
                trim_process_csv(structure_name,file_name)
            if index == 1:
                test_area_structure3 = ["279.608,42.381,294.908,579.411"]
                df_structure3 = tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages=[index+1],area=test_area_structure3,silent=True,guess=False)
                structure_name=df_structure3[0].columns[0]
                df2=tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages=[index + 1],area=test_area1,columns=column_values1,silent=True,guess=False)    
                file_name='file2.csv'
                df2[0].to_csv(file_name)
                trim_process_csv(structure_name,file_name)
                break
        print("Done")
    except Exception as e:
        print(e)    
print('dfv')      
read_pdf()

     