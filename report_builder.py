import os
from datetime import datetime

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.shared import Inches, Pt
from windrose import WindroseAxes

import weather_station


def make_report(
    station: weather_station.WeatherStation,
    date_from: datetime,
    date_to: datetime,
    snow: bool,
    metel: bool,
    rose_type: int,
    doc_name: str,
):
    """Основная функция обработки данных и генерации отчетов"""
    print(f"Обработка файла: {station.csv_path}")
    print(f"Период: с {date_from} по {date_to}")
    print(f"Условия: снег={snow}, ветер≥3м/с={metel}")

    metadata = station.get_metadata()

    data = station.get_data_for_rose_of_wind(
        date_from,
        date_to,
        snow,
        metel,
    )

    # Проверка, остались ли данные после фильтрации
    if data.empty:
        print("Нет данных, соответствующих критериям фильтрации")
        return

    # 1. Абсолютная таблица
    abs_pivot = _create_absolute_pivot(data)

    # 2. Процентная таблица
    total = abs_pivot.sum().sum()

    rel_pivot = _create_percentage_pivot(abs_pivot, total_count=total)

    csv_dir = os.path.dirname(station.csv_path)
    image_full_path = os.path.join(csv_dir, f"{metadata[3]}.jpg")

    # Построение розы ветров
    _draw_and_save_wind_rose_plot(
        data,
        rose_type,
        image_full_path,
        **{"total_count": total},
    )

    # Генерация Word документа
    _render_to_word(
        date_from,
        date_to,
        snow,
        metel,
        metadata[1],
        rel_pivot,
        image_full_path,
        total,
        doc_name,
    )


def _draw_and_save_wind_rose_plot(data, rose_name, save_to, **kwargs):
    """
    Build Wind Rose Figure and save it to file.

    Kwargs:

    - total_count: int
    """

    total_count = kwargs.get("total_count", None)

    print("Построение розы ветров...")

    # Словарь для преобразования направлений ветра в градусы
    w_dir = {
        "N": 0,
        "NNE": 22.5,
        "NE": 45,
        "ENE": 67.5,
        "E": 90,
        "ESE": 112.5,
        "SE": 135,
        "SSE": 157.5,
        "S": 180,
        "SSW": 202.5,
        "SW": 225,
        "WSW": 247.5,
        "W": 270,
        "WNW": 292.5,
        "NW": 315,
        "NNW": 337.5,
    }

    # Преобразование буквенных обозначений в градусы
    data = data.copy()

    data["wd"] = data["wd"].apply(lambda x: w_dir[x] if x in w_dir else np.nan)

    data = data.dropna()

    if len(data) == 0:
        print("Нет данных для построения розы ветров после преобразования направлений")
        return

    # Создание оси для розы ветров
    fig = plt.figure(figsize=(10, 8))

    ax = WindroseAxes.from_ax(fig=fig)

    # Определяем бины для скорости ветра
    # Используем бины от 0 до максимальной скорости с шагом 1 м/с
    max_ws = int(data["ws"].max()) + 1
    bins = np.arange(0, max_ws + 1, 1)

    print(
        f"Диапазон скоростей: от {data['ws'].min():.1f} до {data['ws'].max():.1f} м/с",
    )
    print(f"Бины: {bins}")

    # Построение графика в зависимости от типа
    if rose_name == 0:
        # Черно-белая роза ветров
        ax.contourf(data["wd"], data["ws"], bins=bins, cmap=plt.cm.Greys)
        ax.contour(data["wd"], data["ws"], bins=bins, colors="k", lw=1)
    elif rose_name == 1:
        # Цветная роза ветров с контуром
        ax.contourf(data["wd"], data["ws"], bins=bins, cmap=plt.cm.viridis)
        ax.contour(data["wd"], data["ws"], bins=bins, colors="k", lw=1)
    else:
        # Только контур (по умолчанию)
        ax.contour(data["wd"], data["ws"], bins=bins, colors="k", lw=2)

    # Настройка отображения
    ax.set_xticklabels(
        ["В", "СВ", "С", "СЗ", "З", "ЮЗ", "Ю", "ЮВ"], fontsize=16, fontweight="bold"
    )
    ax.set_theta_zero_location("E")

    # Добавляем легенду с процентами
    ax.set_legend(
        title="Скорость ветра, м/с", loc="center left", bbox_to_anchor=(1.1, 0.5)
    )

    # ВЫЧИСЛЯЕМ РЕАЛЬНЫЕ ПРОЦЕНТЫ ИЗ ДАННЫХ ДЛЯ КРУГОВ

    # 1. Сначала получаем распределение данных по направлениям
    direction_counts = {}
    for direction in w_dir.values():
        # Считаем количество записей для каждого направления
        direction_data = data[data["wd"] == direction]
        direction_counts[direction] = len(direction_data)

    # 2. Находим максимальный процент среди всех направлений
    if total_count is None:
        total_count = len(data)

    max_percentage = 0
    for direction, count in direction_counts.items():
        percentage = (count / total_count * 100) if total_count > 0 else 0
        if percentage > max_percentage:
            max_percentage = percentage

    print(f"Максимальный процент среди направлений: {max_percentage:.1f}%")
    print(f"Общее количество данных: {total_count}")

    # 3. Настраиваем круги на основе реальных данных
    # Всегда используем 4 кольца

    # Стандартные относительные позиции: 0.25, 0.5, 0.75, 1.0 (25%, 50%, 75%, 100%)
    standard_positions = np.array([0.25, 0.5, 0.75, 1.0])

    # Вычисляем фактические значения процентов на основе максимального
    if max_percentage > 0:
        # Масштабируем стандартные позиции под наши данные
        circle_values = standard_positions * max_percentage
    else:
        # Дефолтные значения, если нет данных
        circle_values = np.array([5, 10, 15, 20])

    # Округляем значения кругов
    circle_values = np.round(circle_values, 1)

    # Убеждаемся, что последнее кольцо соответствует максимальному проценту
    circle_values[-1] = max_percentage

    # Преобразуем проценты в позиции на графике
    ylim = ax.get_ylim()
    max_radius = ylim[1]

    if max_radius <= 0:
        max_radius = 1.0

    # Позиции кругов пропорциональны стандартным позициям
    circle_positions = standard_positions * max_radius

    # Устанавливаем круги
    ax.set_yticks(circle_positions)

    # Создаем метки для кругов
    circle_labels = []
    for value in circle_values:
        if value < 0.1:
            circle_labels.append(f"{value:.2f}%")
        elif value < 1:
            circle_labels.append(f"{value:.1f}%")
        else:
            if value.is_integer():
                circle_labels.append(f"{int(value)}%")
            else:
                circle_labels.append(f"{value:.1f}%")

    ax.set_yticklabels(circle_labels, fontsize=11, fontweight="bold")

    # Устанавливаем круги
    ax.set_yticks(circle_positions)

    # Создаем метки для кругов - проценты
    circle_labels = [f"{value:.1f}%" for value in circle_values]
    ax.set_yticklabels(circle_labels, fontsize=11, fontweight="bold")

    # Сохранение графика
    plt.tight_layout()
    plt.savefig(save_to, dpi=150, bbox_inches="tight")
    plt.close()

    print(f"Роза ветров сохранена как {save_to}")
    print(f"Круги отображают проценты: {', '.join(circle_labels)}")


def _create_absolute_pivot(data: pd.DataFrame):
    res = data.groupby(["ws", "wd"]).size().reset_index(name="count")

    pivot_abs = pd.pivot_table(
        res, values="count", index="ws", columns="wd", aggfunc="sum", fill_value=0
    )

    return pivot_abs


def _create_percentage_pivot(abs_pivot, total_count=None):
    if total_count is None:
        total_count = abs_pivot.sum().sum()

    # 1. Проценты и округление
    pct_table = abs_pivot / total_count * 100

    # 2. Итог по направлениям → новая строка
    direction_totals_pct = pct_table.sum(axis=0)
    pct_table.loc["Всего,%"] = direction_totals_pct

    # 3. Итог по скоростям → новый столбец
    speed_totals_pct = pct_table.iloc[:-1].sum(axis=1)
    pct_table["Всего,%"] = 0.0  # создаём столбец
    # Присваиваем значения только для строк, которые НЕ являются 'Всего,%'
    pct_table.loc[speed_totals_pct.index, "Всего,%"] = speed_totals_pct

    # 4. Пересечение итогов
    total_sum = pct_table.iloc[:-1, :-1].sum().sum()
    pct_table.loc["Всего,%", "Всего,%"] = total_sum

    return pct_table


# Константы направлений
DIRECTIONS_ORDER = [
    "N",
    "NNE",
    "NE",
    "ENE",
    "E",
    "ESE",
    "SE",
    "SSE",
    "S",
    "SSW",
    "SW",
    "WSW",
    "W",
    "WNW",
    "NW",
    "NNW",
]

DIRECTION_LABELS = [
    "С",
    "ССВ",
    "СВ",
    "СВС",
    "В",
    "ВЮВ",
    "ЮВ",
    "ЮЮВ",
    "Ю",
    "ЮЮЗ",
    "ЮЗ",
    "ЗЮЗ",
    "З",
    "ЗСЗ",
    "СЗ",
    "ССЗ",
]

assert len(DIRECTIONS_ORDER) == len(DIRECTION_LABELS), (
    "Несоответствие количества направлений и меток"
)

DIR_MAPPING = dict(zip(DIRECTIONS_ORDER, DIRECTION_LABELS))

TABLE_HEADERS = ["Скорость ветра, м/с / Направление"] + \
    DIRECTION_LABELS + ["Всего,%"]


def _format_value(value: float) -> str:
    """Форматирует значение для отображения: 0.0 если < 0.05, иначе с 1 знаком."""
    if pd.isna(value) or abs(value) < 0.05:
        return "0.0"
    return f"{value:.3f}"


def _get_sorted_speeds(index: pd.Index) -> list:
    """Сортирует индексы скоростей по числовому значению, нечисловые — в конец."""

    def key(x):
        try:
            return (0, float(x))
        except (ValueError, TypeError):
            return (1, str(x))

    return sorted([idx for idx in index if idx != "Всего,%"], key=key)


def _render_to_word(
    date_from,
    date_to,
    snow,
    metel,
    case_city,
    rel_pivot: pd.DataFrame,
    rose_of_wind_img_path: str,
    total_count: int,
    doc_name: str,
):
    # Формирование описания условий
    if snow:
        sn = "Осадки в виде снега, "
    else:
        sn = "Независимо от осадков,"

    if metel:
        sn = sn + "ветер 3 и более м/с. "
    else:
        sn = sn + "независимо от скорости ветра."

    region = f"{case_city}. Данные с {date_from} по {date_to}\n{sn}"

    print("Создание Word документа...")

    # === 1. Создание документа и стилей ===
    document = Document()
    style = document.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(14)

    # Заголовок
    p = document.add_paragraph()
    p.alignment = WD_TABLE_ALIGNMENT.CENTER
    p.add_run(f"Роза ветров в {region}")

    # Общее количество записей
    p2 = document.add_paragraph()
    p2.alignment = WD_TABLE_ALIGNMENT.CENTER
    p2.add_run(f"Всего обработано записей: {total_count}")

    # === 2. Добавление изображения ===
    if os.path.exists(rose_of_wind_img_path):
        document.add_picture(rose_of_wind_img_path, width=Inches(6))
    else:
        document.add_paragraph(
            f"Изображение не найдено: {rose_of_wind_img_path}",
        )

    # === 3. Альбомная ориентация для таблицы ===
    current_section = document.sections[-1]
    new_section = document.add_section(WD_SECTION.NEW_PAGE)
    new_section.orientation = WD_ORIENT.LANDSCAPE
    new_section.page_width = current_section.page_height
    new_section.page_height = current_section.page_width

    # Уменьшаем шрифт для таблицы
    style.font.size = Pt(9)

    # === 4. Создание таблицы ===
    n_cols = len(TABLE_HEADERS)
    table = document.add_table(rows=1, cols=n_cols)
    table.style = "Table Grid"

    # Заголовки
    for j, header in enumerate(TABLE_HEADERS):
        cell = table.rows[0].cells[j]
        cell.text = header
        cell.paragraphs[0].alignment = WD_TABLE_ALIGNMENT.CENTER
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # === 5. Подготовка данных ===
    # Убедимся, что все нужные направления есть в данных (даже если 0)
    available_dirs = [d for d in DIRECTIONS_ORDER if d in rel_pivot.columns]
    missing_dirs = set(DIRECTIONS_ORDER) - set(available_dirs)
    for d in missing_dirs:
        rel_pivot[d] = 0.0

    # Сортируем строки: сначала скорости, потом 'Всего,%'
    speed_rows = _get_sorted_speeds(rel_pivot.index)
    all_rows = speed_rows + \
        (["Всего,%"] if "Всего,%" in rel_pivot.index else [])

    # === 6. Заполнение строк таблицы ===
    for row_label in all_rows:
        row_cells = table.add_row().cells

        # Первая ячейка: метка строки
        if row_label == "Всего,%":
            row_cells[0].text = "Всего,%"
        else:
            try:
                ws_int = int(float(row_label))
                row_cells[0].text = str(ws_int)
            except (ValueError, TypeError):
                row_cells[0].text = str(row_label)
        row_cells[0].paragraphs[0].alignment = WD_TABLE_ALIGNMENT.CENTER

        # Направления (столбцы 1–16)
        for j, dir_code in enumerate(DIRECTIONS_ORDER):
            value = (
                rel_pivot.loc[row_label, dir_code]
                if dir_code in rel_pivot.columns
                else 0.0
            )
            row_cells[j + 1].text = _format_value(value)
            row_cells[j + 1].paragraphs[0].alignment = WD_TABLE_ALIGNMENT.RIGHT

        # Последний столбец: 'Всего,%'
        total_col_val = (
            rel_pivot.loc[row_label, "Всего,%"]
            if "Всего,%" in rel_pivot.columns
            else 0.0
        )
        row_cells[-1].text = _format_value(total_col_val)
        row_cells[-1].paragraphs[0].alignment = WD_TABLE_ALIGNMENT.RIGHT

    # === 7. Строка проверки суммы (опционально) ===
    if "Всего,%" in rel_pivot.index:
        direction_sum = sum(
            rel_pivot.loc["Всего,%", d]
            for d in DIRECTIONS_ORDER
            if d in rel_pivot.columns
        )
        check_row = table.add_row().cells
        check_row[0].text = "Проверка суммы:"
        msg_cell = check_row[1]
        msg_cell.text = f"Сумма по направлениям: {direction_sum:.1f}%"
        # Объединяем ячейки 1–16
        for j in range(2, 17):
            msg_cell.merge(check_row[j])
        msg_cell.paragraphs[0].alignment = WD_TABLE_ALIGNMENT.CENTER

    # === 8. Сохранение ===
    try:
        document.save(doc_name)
        print(f"Документ сохранён: {doc_name}")
    except Exception as e:
        print(f"Ошибка при сохранении документа: {e}")
