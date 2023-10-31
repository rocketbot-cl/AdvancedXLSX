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
import datetime
import traceback
import dateutil.parser


base_path = tmp_global_obj["basepath"]
cur_path = base_path + 'modules' + os.sep + 'AdvancedXLSX' + os.sep + 'libs' + os.sep
cur_path_x64 = os.path.join(cur_path, 'Windows' + os.sep +  'x64' + os.sep)
cur_path_x86 = os.path.join(cur_path, 'Windows' + os.sep +  'x86' + os.sep)

if sys.maxsize > 2**32 and cur_path_x64 not in sys.path:
    sys.path.append(cur_path_x64)
elif sys.maxsize <= 2**32 and cur_path_x86 not in sys.path:
    sys.path.append(cur_path_x86)

import openpyxl
from openpyxl.utils.cell import column_index_from_string
from advanced_xlsx import AdvancedXlsx
from whichOperation import whichOperation


module = GetParams("module")

try:
    excel = GetGlobals("xlsx")
except:
    excel = GetGlobals("xls")

if excel.actual_id in excel.file_:
    if "workbook" in excel.file_[excel.actual_id]:
        wb = excel.file_[excel.actual_id]["workbook"]
        advanced_xlsx = AdvancedXlsx(wb)

if module == "open_xls":
    path = GetParams("path")
    id_ = GetParams("id")
    var_ = GetParams("var_")
    col = GetParams("col")
    
    if col:
        col = col.split(",")
    
    if not id_:
        id_ = "default"
    
    try:
        advanced_xlsx = AdvancedXlsx()
        wb = advanced_xlsx.open_xls(path, col)
        excel.actual_id = id_
        excel.file_[excel.actual_id]= {
            "workbook": wb,
            "sheet": wb.active
        }
        SetVar(var_, True)
    except Exception as e:
        print("Traceback: ", traceback.format_exc())
        SetVar(var_, False)
        PrintException()

if module == "xls_to_xlsx":
    xls_path = GetParams("xls_path")
    xlsx_path = GetParams("xlsx_path")
    try:
        advanced_xlsx = AdvancedXlsx()       
        wb = advanced_xlsx.open_xls(xls_path)
        wb.save(filename=xlsx_path)
    except Exception as e:
        print("Traceback: ", traceback.format_exc())
        PrintException()
        raise e
        
if module == "convert_to_csv":
    csv_path = GetParams("csv_path")
    delimiter = GetParams("delimiter") or ","
    result = GetParams("var_")
    
    try:
        
        advanced_xlsx.convert_to_csv(csv_path, delimiter)
        SetVar(result, True)
    except Exception as e:
        print("Traceback: ", traceback.format_exc())
        PrintException()
        SetVar(result, False)
        raise e

if module == "readRange":
    
    sheet = GetParams("sheet_name")
    range_ = GetParams("range")
    var = GetParams("var")
    
    try:
        
        data = advanced_xlsx.read_range(sheet, range_)
        
        if not isinstance(data, tuple):
            data = data.value.isoformat() if isinstance(data.value, datetime.datetime) or isinstance(data.value, datetime.time) else data.value
        # data = [[x.value for x in i] for i in data if isinstance(i, tuple)]
        try:
            data = [[x.value.isoformat() if isinstance(x.value, datetime.datetime) or isinstance(x.value, datetime.time) else x.value for x in i] for i in data if isinstance(i, tuple)]
        except:
            data = [i.value.isoformat() if isinstance(i.value, datetime.datetime) or isinstance(i.value, datetime.time) else i.value for i in data if isinstance(i, tuple)]
        SetVar(var, data)
    
    except Exception as e:
        SetVar(var, False)
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e

if module == "format_cell":

    try:
        sheet_ = GetParams("sheet")
        range_ = GetParams("range")
        # col = "True"
        # row = "False"
        format_code = GetParams("format")
        var_ = GetParams("var_")
        horizontal = GetParams("horizontal")
        vertical = GetParams("vertical")

        import re
        
        if not format_code or format_code == "":
            format_code = 0
        else:
            try:
                int(format_code)
            except:
                raise Exception("Code must be an integer from the Built in Formats.")

        if range_:
            regex = r"([a-zA-Z]*)([0-9]*):([a-zA-Z]*)([0-9]*)"
            matches = re.match(regex, range_).groups()
            rows = [(int(matches[1]) if matches[1] != "" else ""), (int(matches[3]) if matches[3] != "" else "")]
            cols = [matches[0], matches[2]]
        
        # if col:
        #     col = eval(col)
        #     if col == True:
        #         advanced_xlsx.change_format_col(sheet_, cols, format_code, horizontal, vertical)        
        # if row:
        #     row = eval(row)
        #     if row == True:
        #         advanced_xlsx.change_format_row(sheet_, rows, format_code, horizontal, vertical)
        
        # if not col and not row:
        advanced_xlsx.change_format(sheet_, range_, format_code, horizontal, vertical)
     
        SetVar(var_, True)
    except Exception as e:
        SetVar(var_, False)
        print("Traceback: ", traceback.format_exc())
        PrintException()
        raise e

if module == "createSheet":
    try:
        name = GetParams("name")
        advanced_xlsx.new_sheet(name)
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e

if module == "removeSheet":
    try:
        name = GetParams("name")

        advanced_xlsx.del_sheet(wb[name])
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e

if module == "countRange":
    
    sheet = GetParams("sheet_name")
    range_ = GetParams("range")
    row_var = GetParams("row")
    col_var = GetParams("column")
    
    try:
        advanced_xlsx.change_sheet(sheet)
        
        col_length, row_length = advanced_xlsx.count_by_range(range_)
        
        if row_var:
            SetVar(col_var, row_length)
        if col_var:
            SetVar(row_var, col_length)
    
    except Exception as e:
        if row_var:
            SetVar(col_var, False)
        if col_var:
            SetVar(row_var, False)
        
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e

if module == "get_by_filter":
    advanced_xlsx.get_cells_by_range()

if module == "advanceFilter": 

    sheet = GetParams("sheetName")

    if not sheet:
        ws = wb.active
    else:
        ws = wb.get_sheet_by_name(sheet)

    # ws = advanced_xlsx.change_sheet(sheet)

    userFilters = GetParams("userFilters")
    userFilters = eval(userFilters)
    filtros = userFilters
    whereToStoreResult = GetParams("whereToStoreResult")

    variableConTodo = []
    firstFilter = filtros[0]
    firstFilterSplited = firstFilter.split(' ')
    tipo = None

    if (len(firstFilterSplited) == 2):
        tipo = "re"
        firstFilterSplited.append('')
    elif (len(firstFilterSplited) == 3):
        tipo = "common"

    if (tipo == "common"):
        firstFilterSplited[2] = firstFilterSplited[2].replace('%', ' ')
        firstFilterSplited[2] = firstFilterSplited[2].replace('\'', '')

    try:
        # Parses numbers
        if "/" not in firstFilterSplited[2] and "-" not in firstFilterSplited[2]:
            firstFilterSplited[2] = eval(firstFilterSplited[2])
        # Parses dates
        else:
            firstFilterSplited[2] = dateutil.parser.parse(firstFilterSplited[2], dayfirst=True).isoformat()
            firstFilterSplited[2] = advanced_xlsx.get_excel_date(firstFilterSplited[2])
        # Do nothing
    except:
        pass

    for index, row in enumerate (ws.iter_rows()):
        columna = column_index_from_string(firstFilterSplited[0])
        columna -= 1
        cellValue = row[columna].value
        
        
        if isinstance(cellValue, datetime.datetime):
            row[columna].value = cellValue.isoformat()
            cellValue = advanced_xlsx.get_excel_date(cellValue.isoformat())
        
        if (isinstance(cellValue, str) and tipo == "common" and firstFilterSplited[1] != "=="):
            try:
                # Checks if the string is a date and parses it
                cellValue = dateutil.parser.parse(firstFilterSplited[2], dayfirst=True).isoformat()
                cellValue = advanced_xlsx.get_excel_date(cellValue)
            except:
                continue
        #This try/except is to catch the exception when cell values are "NoneType" and can't be compared with "<" or ">"
        try:
            if (whichOperation(cellValue, firstFilterSplited[1], firstFilterSplited[2], tipo)):
                variableConTodo.append([{"row" : f"{index}", "data" : row}])
        except:
            continue

    count = 0
    variableConCasiTodo = []
    variableFinal = variableConTodo
    if (len(filtros) > 1):
        for filtro in filtros:
            if (count == 0):
                count += 1
            else:
                filtroSplited = filtro.split(' ')
                if (len(filtroSplited) == 2):
                    tipo = "re"
                    filtroSplited.append(0)
                elif (len(filtroSplited) == 3):
                    tipo = "common"
                for index, row in enumerate (variableFinal):
                    columna = column_index_from_string(filtroSplited[0])
                    columna -= 1
                    xlRow = row[0]["data"]
                    realRow = row[0]["row"]
                    cellValue = xlRow[columna].value
                    
                    if isinstance(cellValue, datetime.datetime):
                        row[columna].value = cellValue.isoformat()
                        cellValue = advanced_xlsx.get_excel_date(cellValue.isoformat())
                    
                    if (isinstance(cellValue, str) and tipo == "common" and firstFilterSplited[1] != "=="):
                        try:
                            # Checks if the string is a date and parses it
                            cellValue = dateutil.parser.parse(firstFilterSplited[2], dayfirst=True).isoformat()
                            cellValue = advanced_xlsx.get_excel_date(cellValue)
                        except:
                            continue
                    #This try/except is to catch the exception when cell values are "NoneType" and can't be compared with "<" or ">"
                    try:
                        if (whichOperation(cellValue, filtroSplited[1], filtroSplited[2], tipo)):
                            variableConCasiTodo.append([{"row": f"{realRow}", "data" : xlRow}])
                    except:
                        continue
                    
                variableFinal = variableConCasiTodo
                variableConCasiTodo = []

    rowFake = None
    provisionaryArray = []
    variableSinDetail = []
    for row in variableFinal:
        cada = row[0]["data"]
        for columna in cada:
            valor = columna.value
            if (valor == None):
                valor = ''
            provisionaryArray.append(valor)
            rowFake = eval(row[0]["row"])
        row[0]["row"] = str(int(row[0]["row"]) + 1)
        row[0]["data"] = provisionaryArray
        variableSinDetail.append(provisionaryArray)
        provisionaryArray = []
    variableConDetail = []

    for i in variableFinal:
        variableConDetail.append(i[0])

    detailedResult = GetParams("detailedResult")

    if (detailedResult == "True"):
        SetVar(whereToStoreResult, variableConDetail)
    else:
        SetVar(whereToStoreResult, variableSinDetail)

if module == "delete_cell":
    sheet = GetParams("sheet_name")
    row_var = GetParams("row")
    col_var = GetParams("column")
    
    try:
        advanced_xlsx.change_sheet(sheet)

        if row_var:
            advanced_xlsx.delete_rows(row_var)

        if col_var:
            advanced_xlsx.delete_columns(col_var)

    except Exception as e:
        print(traceback.print_exc())
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e


if module == "insert_cell":
    sheet = GetParams("sheet_name")
    row_var = GetParams("row")
    col_var = GetParams("column")
    
    try:
        advanced_xlsx.change_sheet(sheet)
        if row_var:
            advanced_xlsx.insert_rows(row_var)

        if col_var:
            advanced_xlsx.insert_columns(col_var)
        
    except Exception as e:
        print(traceback.print_exc())
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e

if module == "renameSheet": 
    try:
        name = GetParams("name")
        new_name = GetParams("new_name")
        
        # rename sheet with new name
        wb[name].title = new_name

    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e

if module == "insert_image":
    try:
        sheet = GetParams("sheet")
        path = GetParams("path")
        cell = GetParams("cell")
        
        advanced_xlsx.insert_image(sheet, path, cell)
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e