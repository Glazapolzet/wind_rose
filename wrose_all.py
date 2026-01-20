# -*- coding: utf-8 -*-
import pandas as pd

from datetime import date

import wind_plot
import os

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QAction, QToolBar, QMessageBox, QToolTip, QFileDialog, QInputDialog, QWidget,QStatusBar
from PyQt5.QtGui import QPixmap, QFont
import sys
from dataclasses import dataclass


@dataclass
class MeteostationLoaderDialogResult:
    file_path: str
    city_name: str
    city_case: str


class MeteostationLoaderDialog:
    def __init__(self, centralwidget: QWidget):
        self._centralwidget = centralwidget

    def query_data_safe(
            self,
    ) -> tuple[MeteostationLoaderDialogResult | None, bool]:
        try:
            data = self.query_data()

            return data, True

        except BaseException as e:
            QMessageBox.critical(
                self._centralwidget,
                "Ошибка",
                f"Введенные даныне не верны: {str(e)}"
            )

        return None, False

    def query_data(self) -> MeteostationLoaderDialogResult:
        """Загрузка CSV файла с данными метеостанции"""
        file_path, ok = QFileDialog.getOpenFileName(
            self._centralwidget,
            "Выберите CSV файл с данными метеостанции",
            "",
            "CSV files (*.csv);;All files (*.*)"
        )

        if not ok or not file_path:
            raise ValueError("bad filepath")

        city_name, ok = QInputDialog.getText(
            self._centralwidget,
            "Название метеостанции",
            "Введите название города (в именительном падеже):"
        )

        if not ok or not city_name:
            raise ValueError("bad city name")

        # Запрос названия в предложном падеже
        case_city, ok = QInputDialog.getText(
            self._centralwidget,
            "Название в предложном падеже",
            f"Введите название города в предложном падеже (например, 'в {city_name}'):",
            text=f"{city_name}"
        )

        if not ok or not case_city:
            raise ValueError("bad case city")

        return MeteostationLoaderDialogResult(**{
            "file_path": file_path,
            "city_name": city_name,
            "city_case": case_city,
        })


@dataclass
class LoadMeteostationQuery:
    name: str
    path: str


def load_default_stations(
        queries: list[LoadMeteostationQuery],
) -> list[wind_plot.WeatherStation]:
    stations = []

    for query in queries:
        if not os.path.exists(query.path):
            continue

        try:
            station = wind_plot.WeatherStation(query.path, query.name)

            stations.append(station)

            print(f"Загружена станция по умолчанию: {query.name}")

        except Exception as e:
            print(f"Ошибка при загрузке станции {query.name}: {e}")

    return stations


class MeteostationsIndex:
    def __init__(
        self,
        centralwidget: QWidget,
        statusbar: QStatusBar,
        default_stations: list[wind_plot.WeatherStation]
    ):
        self._centralwidget = centralwidget

        self._statusbar = statusbar

        # TODO: make private
        self._qcombobox = QtWidgets.QComboBox(self._centralwidget)
        self._qcombobox.setGeometry(QtCore.QRect(10, 110, 261, 41))
        self._qcombobox.setStyleSheet("""
            background-color: rgb(255, 255, 255);
            color: rgb(0, 0, 0);
        """)
        self._qcombobox.setObjectName(
            "meteostation",
        )
        self._qcombobox.setToolTip(
            'Выберите метеостанцию из списка загруженных',
        )

        self._stations = default_stations

        if self._stations is None:
            self._stations = []

        self._add_item_dialog = MeteostationLoaderDialog(self._centralwidget)

        self._qcombobox.addItems(
            [item.city_name for item in self._stations],
        )

    def widgets(self) -> list[QWidget]:
        return [self._qcombobox]

    def get_current_selected_station(self) -> wind_plot.WeatherStation | None:
        selected_index = self._qcombobox.currentIndex()

        if selected_index < 0 or selected_index > len(self._stations):
            return None

        return self._stations[selected_index]

    def remove_current_selected_station(self):
        selected_index = self._qcombobox.currentIndex()

        if selected_index < 0 or selected_index > len(self._stations):
            return

        self._stations.pop(selected_index)
        self._qcombobox.removeItem(selected_index)

    def get_current_stations(self) -> list[wind_plot.WeatherStation]:
        return self._stations

    def handle_add_station(self):
        data, ok = self._add_item_dialog.query_data_safe()
        if not ok or not data:
            return

        try:
            # Создание объекта метеостанции
            station = wind_plot.WeatherStation(
                data.file_path,
                data.city_name,
                data.city_case,
            )

            self._stations.append(station)

            self._qcombobox.addItem(data.city_name)

            QMessageBox.information(
                self._centralwidget,
                "Станция загружена",
                f"Метеостанция '{data.city_name}' успешно загружена."
            )

            # Обновление статуса
            self._statusbar.showMessage(
                f"Загружена станция: {data.city_name}", 3000,
            )

        except Exception as e:
            QMessageBox.critical(
                self._centralwidget,
                "Ошибка",
                f"Не удалось загрузить файл: {str(e)}"
            )

    def handle_remove_station(self):
        """Удаление выбранной метеостанции из списка"""
        if not self._stations:
            QMessageBox.warning(
                self._centralwidget,
                "Предупреждение",
                "Список станций пуст.",
            )

            return

        station = self.get_current_selected_station()

        if not station:
            return

        reply = QMessageBox.question(
            self._centralwidget,
            "Подтверждение",
            f"Вы уверены, что хотите удалить станцию '{station.city_name}'?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.remove_current_selected_station()
            self._statusbar.showMessage("Станция удалена", 3000)


def make_file_name_from_station(station: wind_plot.WeatherStation) -> str:
    return f'Роза ветров в {station.city_name}. График и таблица.docx'

@dataclass
class RoseOfWindFormResult:
    date_from: date
    date_to: date
    has_snow: bool
    has_wind_over_3m_per_s: bool
    meteostation: wind_plot.WeatherStation
    type_of_rose: int
    type_of_rose_str: str


class RoseOfWindForm:
    def __init__(
        self,
        centralwidget: QWidget,
        meteostations: MeteostationsIndex,
    ):
        self._centralwidget = centralwidget

        self.meteostation = meteostations

        # Выбор типа розы ветров
        self._rose_of_wind_type_widget = QtWidgets.QComboBox(
            self._centralwidget,
        )
        self._rose_of_wind_type_widget.setGeometry(
            QtCore.QRect(10, 170, 261, 41),
        )
        self._rose_of_wind_type_widget.setStyleSheet("""
                    background-color: rgb(255, 255, 255);
                    color: rgb(0, 0, 0);
                """)
        self._rose_of_wind_type_widget.setObjectName("comboBox")
        self._rose_of_wind_type_widget.addItem(
            "Роза ветров цветная",
        )
        self._rose_of_wind_type_widget.addItem(
            "Роза ветров черно-белая",
        )
        self._rose_of_wind_type_widget.addItem(
            "Роза ветров контур",
        )
        self._rose_of_wind_type_widget.setToolTip(
            'Выберите представление розы ветров',
        )

        # Чекбоксы для условий
        self._has_snow_widget = QtWidgets.QCheckBox(
            self._centralwidget,
        )
        self._has_snow_widget.setGeometry(
            QtCore.QRect(300, 330, 81, 21)
                                          )
        self._has_snow_widget.setObjectName(
            "checkBox",
        )
        self._has_snow_widget.setText(
            "Снег",
        )
        self._has_snow_widget.setToolTip(
            'Учитывать только случаи с осадками в виде снега',
        )

        self._has_wind_over_3m_per_s_widget = QtWidgets.QCheckBox(
            self._centralwidget,
        )
        self._has_wind_over_3m_per_s_widget.setGeometry(
            QtCore.QRect(300, 370, 141, 20),
        )
        self._has_wind_over_3m_per_s_widget.setObjectName(
            "checkBox_2",
        )
        self._has_wind_over_3m_per_s_widget.setText(
            "Ветер ≥3 м/с",
        )
        self._has_wind_over_3m_per_s_widget.setToolTip(
            'Учитывать только случаи со скоростью ветра 3 м/с и более',
        )

        # Начальная дата

        self._date_from_label = QtWidgets.QLabel(
            self._centralwidget,
        )
        self._date_from_label.setGeometry(
            QtCore.QRect(280, 440, 191, 31),
        )
        self._date_from_label.setStyleSheet("""
            background-color: transparent;
            color: rgb(0, 0, 0);
            font-weight: bold;
            font-size: 14px;
        """)
        self._date_from_label.setObjectName("label")
        self._date_from_label.setText("Выберите начальную дату")

        self._date_from_widget = QtWidgets.QDateEdit(self._centralwidget)
        self._date_from_widget.setGeometry(QtCore.QRect(280, 470, 194, 22))
        self._date_from_widget.setStyleSheet("""
            background-color: rgb(255, 255, 255);
            color: rgb(0, 0, 0);
        """)
        self._date_from_widget.setObjectName("n_data")
        self._date_from_widget.setMinimumDate(QtCore.QDate(2006, 9, 1))
        self._date_from_widget.setMaximumDate(QtCore.QDate(2024, 9, 1))
        self._date_from_widget.setToolTip('Начальная дата для расчета')

        # Конечная дата
        self._date_to_label = QtWidgets.QLabel(
            self._centralwidget,
        )
        self._date_to_label.setGeometry(
            QtCore.QRect(280, 490, 191, 31),
        )
        self._date_to_label.setStyleSheet("""
            background-color: transparent;
            color: rgb(0, 0, 0);
            font-weight: bold;
            font-size: 14px;
        """)
        self._date_to_label.setObjectName(
            "label_2",
        )
        self._date_to_label.setText(
            "Выберите конечную дату",
        )

        self._date_to_widget = QtWidgets.QDateEdit(self._centralwidget)
        self._date_to_widget.setGeometry(QtCore.QRect(280, 520, 194, 22))
        self._date_to_widget.setStyleSheet("""
            background-color: rgb(255, 255, 255);
            color: rgb(0, 0, 0);
        """)
        self._date_to_widget.setObjectName("k_data")
        self._date_to_widget.setMinimumDate(QtCore.QDate(2010, 8, 31))
        self._date_to_widget.setMaximumDate(QtCore.QDate(2025, 8, 31))
        self._date_to_widget.setToolTip('Конечная дата для расчета')

    def widgets(self) -> list[QWidget]:
        return [
            self._rose_of_wind_type_widget,
            self._has_snow_widget,
            self._has_wind_over_3m_per_s_widget,
            self._date_from_label,
            self._date_from_widget,
            self._date_to_label,
            self._date_to_widget,
        ]

    def query_data(self) -> RoseOfWindFormResult:
        # Проверка наличия станций
        if not self.meteostation.get_current_stations():
            raise ValueError("no uploaded meteostations")

        # Получение выбранной станции
        station = self.meteostation.get_current_selected_station()
        if not station:
            raise ValueError("no meteostation not selected")

        # Получение параметров
        has_snow = self._has_snow_widget.isChecked()  # Снег
        has_wind_over_3m_per_s = self._has_wind_over_3m_per_s_widget.isChecked()  # Ветер ≥3 м/с

        # Получение дат
        date_from = self._date_from_widget.date().toPyDate()
        date_to = self._date_to_widget.date().toPyDate()

        # Проверка дат
        if date_from > date_to:
            raise ValueError("date from if after date to")

        # Словарь для типов розы ветров
        rose_types = {
            "Роза ветров черно-белая": 0,
            "Роза ветров цветная": 1,
            "Роза ветров контур": 2
        }

        # Получение типа розы ветров
        rose_type_str = self._rose_of_wind_type_widget.currentText()

        rose_type = rose_types.get(
            rose_type_str,
            1,
        )

        return RoseOfWindFormResult(**{
            "date_from": date_from,
            "date_to": date_to,
            "has_snow": has_snow,
            "has_wind_over_3m_per_s": has_wind_over_3m_per_s,
            "meteostation": station,
            "type_of_rose": rose_type,
            "type_of_rose_str": rose_type_str,
        })


class Ui_ROSA_VETROV(object):
    QToolTip.setFont(QFont('TimesNewRoman', 10))

    def __init__(self, default_stations: list[wind_plot.WeatherStation]):
        self.stations = default_stations

        if self.stations is None:
            self.stations = []

        self.start_calc = None
        self.centralwidget = None

    def setupUi(self, ROSA_VETROV):
        ROSA_VETROV.setObjectName("ROSA_VETROV")
        ROSA_VETROV.resize(850, 650)  # Немного увеличили размер окна
        ROSA_VETROV.setStyleSheet("background-color: rgb(170, 255, 127);")

        self.centralwidget = QtWidgets.QWidget(ROSA_VETROV)
        self.centralwidget.setObjectName("centralwidget")
        self.centralwidget.setToolTip('Данная программа предназначена для формирования розы ветров.\n'
                                      'Вы можете загружать данные с любых метеостанций.\n'
                                      'Для запуска программы необходимо:\n'
                                      '- загрузить CSV файл с данными\n'
                                      '- выбрать метеостанцию из списка\n'
                                      '- выбрать вид представления розы ветров\n'
                                      '- указать условия фильтрации\n'
                                      '- задать дату начала и окончания выборки данных')

        self.statusbar = QtWidgets.QStatusBar(ROSA_VETROV)
        self.statusbar.setObjectName("statusbar")
        ROSA_VETROV.setStatusBar(self.statusbar)

        self.meteostation = MeteostationsIndex(
            self.centralwidget,
            self.statusbar,
            default_stations,
        )

        self.rose_of_wind_form = RoseOfWindForm(
            self.centralwidget,
            self.meteostation,
        )

        # Кнопка загрузки CSV файла
        self.add_meteostation_from_csv_button = QtWidgets.QPushButton(
            self.centralwidget,
        )
        self.add_meteostation_from_csv_button.setGeometry(
            QtCore.QRect(10, 10, 261, 41),
        )
        self.add_meteostation_from_csv_button.setStyleSheet("""
            background-color: rgb(100, 150, 255);
            color: rgb(255, 255, 255);
            font-weight: bold;
        """)
        self.add_meteostation_from_csv_button.setObjectName(
            "load_csv_button",
        )
        self.add_meteostation_from_csv_button.setText(
            "Загрузить CSV файл",
        )
        self.add_meteostation_from_csv_button.setToolTip(
            'Загрузить CSV файл с данными метеостанции',
        )

        # Кнопка удаления выбранной станции
        self.remove_meteostation_button = QtWidgets.QPushButton(
            self.centralwidget,
        )
        self.remove_meteostation_button.setGeometry(
            QtCore.QRect(10, 60, 261, 41),
        )
        self.remove_meteostation_button.setStyleSheet("""
            background-color: rgb(255, 100, 100);
            color: rgb(255, 255, 255);
            font-weight: bold;
        """)
        self.remove_meteostation_button.setObjectName(
            "remove_station_button")
        self.remove_meteostation_button.setText(
            "Удалить выбранную станцию",
        )
        self.remove_meteostation_button.setToolTip(
            'Удалить выбранную метеостанцию из списка',
        )

        # Кнопка расчета
        self.start_calc = QtWidgets.QPushButton(self.centralwidget)
        self.start_calc.setGeometry(QtCore.QRect(0, 480, 271, 101))
        self.start_calc.setStyleSheet("""
            background-color: rgb(100, 255, 100);
            color: rgb(0, 0, 0);
            font-weight: bold;
            font-size: 14px;
        """)
        self.start_calc.setObjectName("start_calc")
        self.start_calc.setText("РАСЧЕТ")
        self.start_calc.setToolTip('Будет произведен расчет для выбранных условий')

        # Область для графика
        self.graphicsView = QtWidgets.QGraphicsView(self.centralwidget)
        self.graphicsView.setGeometry(QtCore.QRect(390, 50, 331, 301))
        self.graphicsView.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.graphicsView.setObjectName("graphicsView")
        self.graphicsView.setToolTip('Здесь будет представлен график розы ветров')

        # Область для условий расчета
        self.r_cond = QtWidgets.QLabel(self.centralwidget)
        self.r_cond.setGeometry(QtCore.QRect(10, 230, 261, 241))
        self.r_cond.setStyleSheet("""
            background-color: rgb(255, 255, 255);
            color: rgb(0, 0, 0);
            padding: 5px;
        """)
        self.r_cond.setObjectName("r_cond")
        self.r_cond.setWordWrap(True)
        self.r_cond.setToolTip('Здесь будут представлены условия расчета')

        # Кнопка выхода
        self.exit_program = QtWidgets.QPushButton(self.centralwidget)
        self.exit_program.setGeometry(QtCore.QRect(490, 480, 161, 81))
        self.exit_program.setStyleSheet("""
            background-color: rgb(255, 0, 0);
            color: rgb(255, 255, 255);
            font-weight: bold;
        """)
        self.exit_program.setObjectName("exit_program")
        self.exit_program.setText("ВЫХОД")

        # Название розы ветров
        self.rose_name = QtWidgets.QLabel(self.centralwidget)
        self.rose_name.setGeometry(QtCore.QRect(394, 10, 321, 20))
        self.rose_name.setStyleSheet("""
            background-color: rgb(255, 255, 255);
            color: rgb(0, 0, 0);
            font-weight: bold;
            padding: 2px;
        """)
        self.rose_name.setObjectName("rose_name")
        self.rose_name.setToolTip('Название метеостанции и вид розы ветров')

        # Поднимаем все виджеты наверх (чтобы не перекрывались)
        widgets = [
            *self.meteostation.widgets(),
            *self.rose_of_wind_form.widgets(),

            self.add_meteostation_from_csv_button,
            self.remove_meteostation_button,
            self.meteostation._qcombobox,
            self.start_calc,
            self.graphicsView,
            self.r_cond,
            self.exit_program,
            self.rose_name,
        ]

        for widget in widgets:
            widget.raise_()

        ROSA_VETROV.setCentralWidget(self.centralwidget)

        # Меню и статусбар
        self.menubar = QtWidgets.QMenuBar(ROSA_VETROV)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 850, 26))
        self.menubar.setObjectName("menubar")

        # Добавляем меню
        self.menu_file = QtWidgets.QMenu("Файл", self.menubar)
        self.menu_help = QtWidgets.QMenu("Помощь", self.menubar)

        self.menubar.addAction(self.menu_file.menuAction())
        self.menubar.addAction(self.menu_help.menuAction())

        # Добавляем действия в меню
        self.action_load_csv = QtWidgets.QAction("Загрузить CSV", ROSA_VETROV)
        self.action_load_csv.setShortcut("Ctrl+O")
        self.action_exit = QtWidgets.QAction("Выход", ROSA_VETROV)
        self.action_exit.setShortcut("Ctrl+Q")
        self.action_about = QtWidgets.QAction("О программе", ROSA_VETROV)

        self.menu_file.addAction(self.action_load_csv)
        self.menu_file.addSeparator()
        self.menu_file.addAction(self.action_exit)
        self.menu_help.addAction(self.action_about)

        ROSA_VETROV.setMenuBar(self.menubar)

        self.retranslateUi(ROSA_VETROV)
        QtCore.QMetaObject.connectSlotsByName(ROSA_VETROV)

        # Подключение сигналов
        self.connect_handlers()

    def retranslateUi(self, ROSA_VETROV):
        _translate = QtCore.QCoreApplication.translate

        ROSA_VETROV.setWindowTitle(
            _translate("ROSA_VETROV", "Программа построения розы ветров"),
        )

        self.rose_name.setText(
            _translate("ROSA_VETROV", "Выберите метеостанцию и параметры"),
        )

    def connect_handlers(self):
        """Подключение сигналов к слотам"""
        self.start_calc.clicked.connect(self.equal)
        self.exit_program.clicked.connect(app.quit)
        self.add_meteostation_from_csv_button.clicked.connect(self.meteostation.handle_add_station)
        self.remove_meteostation_button.clicked.connect(self.meteostation.handle_remove_station)
        self.action_load_csv.triggered.connect(self.meteostation.handle_add_station)
        self.action_exit.triggered.connect(app.quit)
        self.action_about.triggered.connect(self.show_about)


    def show_about(self):
        """Показать информацию о программе"""
        QMessageBox.about(
            self.centralwidget,
            "О программе",
            "Программа построения розы ветров\n\n"
            "Версия 2.0\n"
            "Поддерживает загрузку CSV файлов с данными любых метеостанций.\n\n"
            "Для работы программы необходим CSV файл с колонками:\n"
            "- dt_time: дата и время\n"
            "- Wind_dir: направление ветра (N, NNE, NE и т.д.)\n"
            "- wind_speed: скорость ветра (м/с)\n"
            "- precipitation: осадки (2 - снег)\n\n"
            "Формат даты: DD.MM.YYYY HH:MM"
        )

    def equal(self):
        """Основная функция расчета"""
        print("Начало расчета...")

        try:
            data = self.rose_of_wind_form.query_data()

            # Обновление названия в интерфейсе
            s = f"{data.meteostation.city_name}: {data.type_of_rose_str}"
            self.rose_name.setText(s)

            save_to = make_file_name_from_station(data.meteostation)

            # Формирование строки условий
            s_r = 'Расчеты проводятся для: \n' + \
                'метеостанции: ' + data.meteostation.city_name + '\n' \
                'c ' + data.date_from.strftime("%d.%m.%Y") + \
                ' по ' + data.date_to.strftime("%d.%m.%Y") + ' \n'

            if data.has_snow:
                s_1 = 'при наличии осадков в виде снега \n'
            else:
                s_1 = 'независимо от осадков \n'

            if data.has_wind_over_3m_per_s:
                s_1 = s_1 + 'и при ветре 3 и более м/с \n'
            else:
                s_1 = s_1 + 'независимо от скорости ветра \n'

            s_r = s_r + s_1

            s_r = s_r + 'Роза ветров и таблица сохранены в файле: \n' + save_to

            # Выполнение расчета
            try:
                wind_plot.obr_file(
                    data.meteostation,
                    data.date_from.strftime("%d.%m.%Y"),
                    data.date_to.strftime("%d.%m.%Y"),
                    data.has_snow,
                    data.has_wind_over_3m_per_s,
                    data.type_of_rose,
                    save_to,
                )

                self.r_cond.setText(s_r)

                # Отображение графика
                metadata = data.meteostation.get_metadata()
                image_path = metadata[3] + '.jpg'

                if os.path.exists(image_path):
                    pix = QPixmap(image_path)
                    pixmap_scaled = pix.scaled(290, 290, QtCore.Qt.KeepAspectRatio)
                    item = QtWidgets.QGraphicsPixmapItem(pixmap_scaled)

                    scene = QtWidgets.QGraphicsScene()
                    scene.addItem(item)
                    self.graphicsView.setScene(scene)
                else:
                    self.r_cond.setText(s_r + '\n\nВнимание: График не найден!')

            except FileNotFoundError as e:
                self.r_cond.setText(f'Ошибка: Файл не найден\n{str(e)}')

            except pd.errors.EmptyDataError:
                self.r_cond.setText('Ошибка: CSV файл пуст или содержит некорректные данные')

            except KeyError as e:
                self.r_cond.setText(f'Ошибка: В CSV файле отсутствуют необходимые колонки\n{str(e)}')

            except Exception as e:
                self.r_cond.setText(f'Ошибка при выполнении расчета:\n{str(e)}')
                print(f"Ошибка: {e}")

        except BaseException as e:
            self.r_cond.setText(str(e))


if __name__ == "__main__":
    default_stations = load_default_stations([
        LoadMeteostationQuery(**{
            "name": "Мценск",
            "path": "archiv_mcensk_05_24.csv",
        }),
        LoadMeteostationQuery(**{
            "name": "Орел",
            "path": "arсhiv_orel_05_24.csv",
        }),
        LoadMeteostationQuery(**{
            "name": "Верховье",
            "path": "arсhiv_verh_05_21.csv",
        }),
    ])

    import sys

    app = QtWidgets.QApplication(sys.argv)

    # Настройка стиля приложения
    app.setStyle('Fusion')

    # Создание главного окна
    rose_wind = QtWidgets.QMainWindow()
    ui = Ui_ROSA_VETROV(default_stations)
    ui.setupUi(rose_wind)

    # Установка иконки (если есть)
    try:
        app.setWindowIcon(QtGui.QIcon('wind_rose_icon.png'))
    except BaseException:
        pass

    # Показ окна
    rose_wind.show()

    # Запуск приложения
    sys.exit(app.exec_())
