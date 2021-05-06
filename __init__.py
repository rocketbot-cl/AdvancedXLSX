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
cur_path = base_path + 'modules' + os.sep + 'AdvancedExcel' + os.sep + 'libs' + os.sep
if cur_path not in sys.path:
    sys.path.append(cur_path)

module = GetParams("module")

try:
    excel = GetGlobals("xlsx")
except:
    excel = GetGlobals("xls")


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
        
        wb = excel.file_[excel.actual_id]["workbook"]
        wb.create_sheet(name)
        

    if module == "countRange":
        wb = excel.file_[excel.actual_id]["workbook"]
        sheet = GetParams("sheet_name")
        range_ = GetParams("range")
        row_var = GetParams("row")
        col_var = GetParams("column")
        
        sheet = wb.get_sheet_by_name("Sheet")
        column = sheet[range_].column
        row = sheet[range_].row
        col_length = len([column for column in sheet.columns][column - 1])
        row_length = len([row for row in sheet.rows][row - 1])
        
        if row_var:
            SetVar(col_var, row_length)

        if col_var:
            SetVar(row_var, col_length)

except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e