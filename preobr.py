import pandas as pd
import numpy as np
import os
import re
import csv
from io import StringIO

# Словарь для преобразования направлений ветра
WIND_DICT = {
    'Ветер, дующий с севера': 'N',
    'Ветер, дующий с северо-северо-востока': 'NNE',
    'Ветер, дующий с северо-востока': 'NE',
    'Ветер, дующий с востоко-северо-востока': 'ENE',
    'Ветер, дующий с востока': 'E',
    'Ветер, дующий с востоко-юго-востока': 'ESE',
    'Ветер, дующий с юго-востока': 'SE',
    'Ветер, дующий с юго-юго-востока': 'SSE',
    'Ветер, дующий с юга': 'S',
    'Ветер, дующий с юго-юго-запада': 'SSW',
    'Ветер, дующий с юго-запада': 'SW',
    'Ветер, дующий с западо-юго-запада': 'WSW',
    'Ветер, дующий с запада': 'W',
    'Ветер, дующий с западо-северо-запада': 'WNW',
    'Ветер, дующий с северо-запада': 'NW',
    'Ветер, дующий с северо-северо-запада': 'NNW',
    'Штиль, безветрие': 'CALM'
}


def _find_header_row_and_time_col(excel_path: str) -> tuple[int, str]:
    """
    Ищет строку-заголовок, содержащую колонку с 'время' или 'time'.
    Возвращает (номер строки заголовка, имя колонки времени).
    """
    # Читаем все строки как строки, без заголовков
    df_raw = pd.read_excel(excel_path, header=None, dtype=str)
    
    # Ищем первую строку, в которой есть колонка, содержащая 'время' или 'time' (регистронезависимо)
    for idx in range(min(20, len(df_raw))):  # проверяем первые 20 строк
        row = df_raw.iloc[idx].dropna().astype(str).str.lower()
        time_cols = row[row.str.contains(r'время|time', na=False)]
        if not time_cols.empty:
            time_col_name = df_raw.iloc[idx][time_cols.index[0]]
            return idx, time_col_name
    
    raise ValueError("Не найдена колонка с временем (содержащая 'время' или 'time') в первых 20 строках")


def _process_common_dataframe(df: pd.DataFrame, time_col: str) -> pd.DataFrame:
    """
    Общая логика обработки датафрейма после чтения из любого источника.
    Выполняет все преобразования, кроме чтения файла.
    """
    # Удаляем лишние колонки (если они есть)
    cols_to_drop = ['N', 'Pa', 'Cm', 'Ch', 'Cl', 'Nh', 'H', "E'", 'E', 'RRR', 'Tn', 'Tx', 'Td', 'tR', 'Tg']
    df = df.drop(columns=[col for col in cols_to_drop if col in df.columns], errors='ignore')
    
    # Переименовываем ключевые колонки
    rename_map = {
        time_col: 'dt_time',
        'sss': 'snow,cm',
        'Ff': 'wind_speed',
        'DD': 'Wind_dir',
        'T': 't_tek'
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})
    
    # Преобразуем направление ветра
    if 'Wind_dir' in df.columns:
        df['Wind_dir'] = df['Wind_dir'].astype(str).str.strip().str.strip('"').map(WIND_DICT).fillna('CALM')
    
    # Обработка осадков — гарантируем наличие колонок W1, W2, WW
    for col in ['W1', 'W2', 'WW']:
        if col not in df.columns:
            df[col] = '0'
        else:
            df[col] = df[col].fillna('0').astype(str).str.strip().str.strip('"')
    
    # Создаём колонку осадков на основе WW/W1/W2
    df['combined'] = (df['WW'] + ' ' + df['W1'] + ' ' + df['W2']).str.lower()
    df['precipitation'] = 0
    
    # Снег (включая "ливневый снег", "слабый снег" и т.п.)
    df.loc[df['combined'].str.contains(r'снег', na=False), 'precipitation'] = 2
    
    # Дождь (только если температура > 0)
    temp_series = pd.to_numeric(df.get('t_tek', pd.Series([np.nan] * len(df))), errors='coerce')
    df.loc[
        (temp_series > 0) &
        (df['combined'].str.contains(r'дождь|ливень', na=False)),
        'precipitation'
    ] = 1
    
    # Удаляем служебные колонки
    df = df.drop(columns=['WW', 'W1', 'W2', 'combined', 'snow,cm'], errors='ignore')
    
    # === КРИТИЧЕСКИ ВАЖНАЯ ПРОВЕРКА: гарантируем наличие обязательных колонок ===
    # Если колонка осадков не была создана (нет данных WW/W1/W2), устанавливаем значение по умолчанию
    if 'precipitation' not in df.columns:
        df['precipitation'] = 0  # Нет данных об осадках = 0 (без осадков)
    
    # Проверяем обязательные колонки для работы приложения
    required_cols = ['dt_time', 'Wind_dir', 'wind_speed', 'precipitation']
    missing_required = [col for col in required_cols if col not in df.columns]
    if missing_required:
        raise ValueError(f"Отсутствуют обязательные колонки после обработки: {missing_required}")
    
    # Формируем финальный набор колонок в правильном порядке
    final_columns = ['dt_time', 't_tek', 'Po', 'P', 'U', 'Wind_dir', 'wind_speed', 'ff10', 'ff3', 'VV', 'precipitation']
    df = df[[col for col in final_columns if col in df.columns]]
    
    return df


def preprocess_excel_to_csv(excel_path: str, city_name: str = None) -> str:
    """
    Обрабатывает Excel-файл метеостанции (в формате rp5.ru) и сохраняет в CSV в требуемом формате.
    Возвращает путь к созданному CSV-файлу.
    """
    # Находим строку заголовка и имя колонки времени
    header_row, time_col = _find_header_row_and_time_col(excel_path)
    
    # Читаем данные с правильным заголовком
    df = pd.read_excel(excel_path, header=header_row)
    
    # Применяем общую логику обработки
    df = _process_common_dataframe(df, time_col)
    
    # Сохраняем CSV в той же папке, что и исходный файл
    output_dir = os.path.dirname(excel_path)
    safe_city = re.sub(r'[^\w\-_]', '_', city_name) if city_name else "station"
    csv_filename = f"{safe_city}_processed.csv"
    csv_path = os.path.join(output_dir, csv_filename)
    
    df.to_csv(csv_path, encoding='cp1251', index=False)
    return csv_path


def preprocess_csv_to_csv(csv_path: str, city_name: str = None) -> str:
    """
    Обрабатывает CSV-файл метеостанции (в формате rp5.ru) и сохраняет в CSV в требуемом формате.
    Возвращает путь к созданному CSV-файлу.
    """
    # Читаем файл с обработкой кодировок
    try:
        with open(csv_path, 'r', encoding='utf-8') as f:
            lines = [line.rstrip('\n\r') for line in f if not line.strip().startswith('#')]
    except UnicodeDecodeError:
        with open(csv_path, 'r', encoding='cp1251') as f:
            lines = [line.rstrip('\n\r') for line in f if not line.strip().startswith('#')]
    
    if not lines:
        raise ValueError("CSV файл содержит только комментарии или пуст")
    
    # Парсим через csv.reader
    reader = csv.reader(lines, delimiter=';', quotechar='"')
    header = next(reader)
    
    # Удаляем пустые колонки из заголовка (часто последняя колонка пустая из-за лишнего ;)
    header = [col.strip() for col in header if col.strip() != '']
    num_header_cols = len(header)
    
    # Собираем данные, корректируя количество колонок
    data = []
    for row in reader:
        # Удаляем пустые элементы в конце (результат лишних разделителей)
        while len(row) > num_header_cols and row and row[-1] == '':
            row.pop()
        # Обрезаем или дополняем до нужного количества колонок
        if len(row) > num_header_cols:
            row = row[:num_header_cols]
        elif len(row) < num_header_cols:
            row.extend([''] * (num_header_cols - len(row)))
        data.append(row)
    
    # Создаём датафрейм
    df = pd.DataFrame(data, columns=header)
    
    # Находим колонку времени
    time_col = None
    for col in df.columns:
        if re.search(r'время|time', str(col), re.IGNORECASE):
            time_col = col
            break
    
    if time_col is None:
        raise ValueError("Не найдена колонка с временем (содержащая 'время' или 'time')")
    
    # Применяем общую логику обработки
    df = _process_common_dataframe(df, time_col)
    
    # Сохраняем результат
    output_dir = os.path.dirname(csv_path)
    safe_city = re.sub(r'[^\w\-_]', '_', city_name) if city_name else "station"
    csv_filename = f"{safe_city}_processed.csv"
    output_path = os.path.join(output_dir, csv_filename)
    
    df.to_csv(output_path, encoding='cp1251', index=False)
    return output_path


def preprocess_file(file_path: str, city_name: str = None) -> str:
    """
    Обрабатывает входной файл (Excel или CSV) и возвращает путь к CSV в правильном формате.
    Проверяет наличие необходимых исходных колонок (для rename_map) перед обработкой.
    """
    file_path = str(file_path)
    ext = file_path.lower()
    
    if ext.endswith(('.xls', '.xlsx')):
        # Проверка наличия необходимых колонок в XLS
        try:
            header_row, time_col = _find_header_row_and_time_col(file_path)
            df_header = pd.read_excel(file_path, header=header_row, nrows=0)
            cols = [str(c).strip() for c in df_header.columns]
        except Exception:
            raise ValueError("XLS файл не содержит необходимых колонок")
        
        # Проверяем наличие ключевых колонок для преобразования
        if 'T' not in cols or 'DD' not in cols or 'Ff' not in cols:
            raise ValueError(
                "XLS файл не содержит необходимых колонок: T (температура), DD (направление ветра), Ff (скорость ветра)")
        
        # Обрабатываем файл
        return preprocess_excel_to_csv(file_path, city_name)
    
    elif ext.endswith('.csv'):
        # Надёжная проверка колонок в CSV
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = [line.rstrip('\n\r') for line in f if not line.strip().startswith('#')]
        except UnicodeDecodeError:
            with open(file_path, 'r', encoding='cp1251') as f:
                lines = [line.rstrip('\n\r') for line in f if not line.strip().startswith('#')]
        
        if not lines:
            raise ValueError("CSV файл пуст или содержит только комментарии")
        
        # Парсим заголовок
        reader = csv.reader(lines, delimiter=';', quotechar='"')
        header = next(reader)
        header = [col.strip() for col in header if col.strip() != '']
        
        # Проверяем наличие ключевых колонок
        has_t = 'T' in header
        has_dd = 'DD' in header
        has_ff = 'Ff' in header
        has_time = any(re.search(r'время|time', col, re.IGNORECASE) for col in header)
        
        if not (has_t and has_dd and has_ff and has_time):
            missing = []
            if not has_t: missing.append('T')
            if not has_dd: missing.append('DD')
            if not has_ff: missing.append('Ff')
            if not has_time: missing.append('время/time')
            raise ValueError(f"CSV файл не содержит необходимых колонок: {', '.join(missing)}")
        
        # Обрабатываем файл
        return preprocess_csv_to_csv(file_path, city_name)
    
    else:
        raise ValueError("Поддерживаются только .csv, .xls, .xlsx файлы")