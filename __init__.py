# coding: utf-8
"""
Base para desarrollo de modulos externos.
Para obtener el modulo/Funcion que se esta llamando:
     GetParams("module")

Para obtener las variables enviadas desde formulario/comando Rocketbot:
    var = GetParams(variable)
    Las "variable" se define en forms del archivo package.json

Para modificar la variable de Rocketbot:
    SetVar(Variable_Rocketbot, "dato")

Para obtener una variable de Rocketbot:
    var = GetVar(Variable_Rocketbot)

Para obtener la Opcion seleccionada:
    opcion = GetParams("option")


Para instalar librerias se debe ingresar por terminal a la carpeta "libs"

    pip install <package> -t .

"""
# Changing the data types of all strings in the module at once
from __future__ import unicode_literals
import os
import sys
import copy
base_path = tmp_global_obj["basepath"]
cur_path = base_path + 'modules' + os.sep + 'AdvancedXlsx' + os.sep + 'libs' + os.sep
if cur_path not in sys.path:
    sys.path.append(cur_path)

from advanced_xlsx import AdvancedXlsx

module = GetParams("module")

try:
    excel = GetGlobals("xlsx")
except:
    excel = GetGlobals("xls")

if "workbook" in excel.file_[excel.actual_id]:
    wb = excel.file_[excel.actual_id]["workbook"]
    advanced_xlsx = AdvancedXlsx(wb)

if module == "GetCell":
    
    try:
        range_ = GetParams("range")
        var_ = GetParams("var_")

        cells =  copy.deepcopy(excel.file_[excel.actual_id]['sheet'][range_])
        
        def Advance_formatCell(cell_, format_):
            print(cell_, format_)
            f_ = ""
            prefix_ = suffix_= ""
            miles = False
            format_ = format_.split(";")[0]
            if isinstance(cell_, (int,float)):
                if format_ == "General" or format_ == "0":
                    return str(cell_)
                
                if format_ == "@":
                    return cell_
                if format_.startswith("$"):
                    prefix_ = prefix_ + "$"
                if '"' in format_:                    
                    prefix_ = prefix_ + format_.split('"')[1]
                if "#,##" in format_ :
                    f_ = f_ +  "2,"
                    miles= True
                
                if "0.0" in format_:
                    count_ = len([ d for d in format_.split("0.")[1] if d and d == '0'])
                    f_ = f_ +  "." + str(count_) + "f"
                
                if  format_ == "0%":
                    f_ = ".0"
                    
                if format_.endswith("%"):                    
                    f_ = f_.replace("f","") + "%"
                
                
                d_ = "{:" + f_ + "}" 
                
                
                data_ =  prefix_ + d_.format(cell_)+ suffix_
                if miles:
                    data_ = data_.replace(",",";").replace(".",",").replace(";",".")
                return data_
            else:
                return cell_

        def Advance_getCells(datas):
            global Advance_formatCell, Advance_getCells
            

            info = []
            
            for data in datas:
                if isinstance(data, tuple):
                    info.append(Advance_getCells(data))
                else:
                    data_ = Advance_formatCell(data.value, data.number_format)
                    
                    if isinstance(data_, (datetime.date, datetime.datetime)):
                        data_ = data_.strftime("%d-%m-%Y")
                    info.append(data_)
            return info
        if isinstance(cells, tuple):
            res = Advance_getCells(cells)
        else:            
            res = Advance_formatCell(cells.value, cells.number_format)
        if not res:
            res = ""
        if isinstance(res, datetime.date):
            res = res.strftime("%d-%m-%Y")
        SetVar(var_, res)
    except:
        PrintException()

try:
    if module == "createSheet":

        name = GetParams("name")
        wb.create_sheet(name)
        

    if module == "countRange":
        
        sheet = GetParams("sheet_name")
        range_ = GetParams("range")
        row_var = GetParams("row")
        col_var = GetParams("column")
        
        advanced_xlsx.change_sheet(sheet)
        
        col_length, row_length = advanced_xlsx.count_by_range(range_)
        
        if row_var:
            SetVar(col_var, row_length)

        if col_var:
            SetVar(row_var, col_length)

    if module == "delete_cell":
        sheet = GetParams("sheet_name")
        row_var = GetParams("row")
        col_var = GetParams("column")
        
        advanced_xlsx.change_sheet(sheet)
        if row_var:
            advanced_xlsx.delete_rows(row_var)

        if col_var:
            advanced_xlsx.delete_columns(col_var)

    if module == "open_xls":
        path = GetParams("path")
        id_ = GetParams("id")

        advanced_xlsx = AdvancedXlsx()
        wb = advanced_xlsx.open_xls(path)
        excel.actual_id = id_
        excel.file_[excel.actual_id]= {
            "workbook": wb,
            "sheet": wb.active
        }

    if module == "get_by_filter":
        advanced_xlsx.get_cells_by_range()

except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e