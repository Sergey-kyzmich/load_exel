import os
import pandas as pd

def extract_text_from_xlsx(file_path) -> dict:
    # Чтение всех листов из файла
    excel_file = pd.ExcelFile(file_path)
    # Создание словаря, где ключ — название листа, значение — соответствующий data_sheetFrame
    sheets_dict = {sheet_name: excel_file.parse(sheet_name) for sheet_name in excel_file.sheet_names}

    data = {}

    for sheet_name, df in sheets_dict.items():

        # Создаем массив строк
        data_sheet = []
        # Добавление списка колонок - 1-я строчка
        data_sheet.append(df.columns.tolist())

        # перебираем колонки
        for column in df.columns:
            # rowid - Номер рассматриваемой строки
            rowid = 1
            # Перебираем все значения в колонке и записываем их
            # в собственный список в массиве
            for item in df[column]:
                # Если индекс строки больше максимального индекса,
                # записанного в массиве
                if rowid > len(data_sheet)-1:
                    data_sheet.append([str(item) if (type(item) != float and item !='nan') else ""])
                else:# Иначе добавляем ячейку к ранее созданному списку ячеек этой строки 
                    data_sheet[rowid] += [str(item) if (type(item) != float and item !='nan') else ""]
                rowid += 1
        
        # Перебираем строки, чтобы найти строку с информацией о структуре столбцов
        for rowid in range(len(data_sheet)):
            # Если строка пуста
            if len(data_sheet[rowid])==0:continue
            # Если строка начинается с "No.", то она содержит
            # информацию о значениях столбца  
            if ("No" in data_sheet[rowid][0] or "№" in data_sheet[rowid][0]) and len(data_sheet[rowid][0])<10:
                # Как только нашлась такая строка - сохраняем rowid и выходим из цикла
                rowid_structure = rowid
                break
        else:# Если не было найдено строки с информацией о структуре столбцов
            rowid_structure = None
        
        # Создаем словарь с ключем - названием листа, значением - словарь значений
        data[sheet_name] = {}

        # Если была найдена строка с типом столбцов, то разделяем строки на 
        # информацию о листе, информацию о структуре столбцов и на остальные строки
        if rowid_structure:
            data[sheet_name]['info'] = data_sheet[:rowid_structure]
            data[sheet_name]['column_structure'] = data_sheet[rowid_structure:rowid_structure+3]
            data[sheet_name]['rows'] = data_sheet[rowid_structure+3:] if rowid_structure+2!= len(data_sheet)-1 else []

        else: 
            # Если не была найдена строка с типом столбцов, 
            # записываем все строки вместе

            # Если лист не пустой
            if len(data_sheet)!=1:
                data[sheet_name]['rows'] = data_sheet 


    return data

def process_file(file_path):
    _, ext = os.path.splitext(file_path)
    if ext.lower() == '.xlsx' or ext.lower() == '.xls':
        return extract_text_from_xlsx(file_path)
    else:
        raise ValueError(f"Unsupported file type {file_path}")