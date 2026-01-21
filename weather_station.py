import os
from dataclasses import dataclass
from datetime import datetime

import numpy as np
import pandas as pd


class WeatherStation:
    """Класс для представления метеостанции и её метаданных"""

    def __init__(self, csv_path, city_name, case_city=None):
        """
        Инициализация метеостанции

        Args:
            csv_path (str): Путь к CSV файлу с данными
            city_name (str): Название города в именительном падеже
            case_city (str, optional): Название в предложном падеже ("в ...")
                                       Если не указано, используется city_name
        """
        self.csv_path = csv_path
        self.city_name = city_name
        self.case_city = case_city or city_name
        self.file_name = os.path.basename(csv_path)

        # # try to get data (validate it)
        # self.get_data()

    def get_metadata(self):
        """Получить метаданные для генерации отчетов"""
        return [
            # [0] для 10-летнего графика
            f"wrose_met {self.city_name} за 10 лет.jpg",
            self.case_city,  # [1] город в предложном падеже
            f"{self.city_name} за период ",  # [2] префикс периода
            f"wrose {self.city_name}",  # [3] базовое имя для сохранения
        ]

    def get_data(self) -> pd.DataFrame:
        data = pd.read_csv(self.csv_path, encoding="cp1251")

        required_columns = [
            "dt_time",
            "Wind_dir",
            "wind_speed",
            "precipitation",
        ]

        missing_columns = [
            col for col in required_columns if col not in data.columns]

        if missing_columns:
            raise ValueError(f"miss columns: {missing_columns}")

        # parse raw fields

        def dateparser(x: str):
            return datetime.strptime(x, "%d.%m.%Y %H:%M").date()

        data["dt_time"] = data["dt_time"].apply(dateparser)

        data["Wind_dir"] = data["Wind_dir"].astype(str).str.strip()

        data["wind_speed"] = pd.to_numeric(data["wind_speed"], errors="coerce")

        data["precipitation"] = pd.to_numeric(
            data["precipitation"],
            errors="coerce",
        )

        return data

    def get_data_in_date_interval(
        self,
        df: datetime,
        dt: datetime,
    ) -> pd.DataFrame:
        data = self.get_data()

        return data.loc[((data["dt_time"] > df) & (data["dt_time"] < dt))]

    def get_data_for_rose_of_wind(
        self,
        date_from: datetime,
        date_to: datetime,
        snow: bool,
        metel: bool,
    ) -> pd.DataFrame:
        """
        Retrieve pd.DataFrame from csv for building rose of wind.

        Fields:

        - wd - wind direction
        - ws - wind speed
        """

        data = self.get_data_in_date_interval(date_from, date_to)

        data = data.copy()

        data = self._remove_calm(data)
        data = self._apply_snow_filter(data, snow)
        data = self._apply_has_wind_over_3m_per_s_filter(data, metel)

        data_wind = pd.DataFrame()

        data_wind["wd"] = data["Wind_dir"]
        data_wind["ws"] = data["wind_speed"]

        return data_wind

    @staticmethod
    def _remove_calm(data: pd.DataFrame) -> pd.DataFrame:
        data = data.copy()

        # Преобразование 'CALM' (штиль) в NaN
        data["Wind_dir"] = data["Wind_dir"].apply(
            lambda x: np.nan if x == "CALM" else x
        )

        data = data[data["Wind_dir"].notna()]

        return data

    @staticmethod
    def _apply_snow_filter(
        data: pd.DataFrame,
        snow: bool,
    ) -> pd.DataFrame:
        # Фильтрация по снегу
        if not snow:
            return data

        data = data.copy()

        # выбор только тех значений, где sn=2 (снег)
        data["precipitation"] = data["precipitation"].apply(
            lambda x: np.nan if x != 2 else x
        )

        data = data[data["precipitation"].notna()]

        return data

    @staticmethod
    def _apply_has_wind_over_3m_per_s_filter(
        data: pd.DataFrame,
        metel: bool,
    ) -> pd.DataFrame:
        # Фильтрация по снегу
        if not metel:
            return data

        data = data.copy()

        data["wind_speed"] = data["wind_speed"].apply(
            lambda x: np.nan if x < 3 else x)

        data = data[data["wind_speed"].notna()]

        return data

    def __str__(self):
        return f"{self.city_name} ({self.file_name})"


@dataclass
class LoadMeteostationQuery:
    name: str
    path: str


def load_default_stations(
        queries: list[LoadMeteostationQuery],
) -> list[WeatherStation]:
    stations = []

    for query in queries:
        if not os.path.exists(query.path):
            continue

        try:
            station = WeatherStation(query.path, query.name)

            stations.append(station)

            print(f"Загружена станция по умолчанию: {query.name}")

        except Exception as e:
            print(f"Ошибка при загрузке станции {query.name}: {e}")

    return stations
