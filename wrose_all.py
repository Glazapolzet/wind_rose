# -*- coding: utf-8 -*-
import pandas as pd

import wind_plot
import os

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QAction, QToolBar, QMessageBox, QToolTip, QFileDialog, QInputDialog
from PyQt5.QtGui import QPixmap, QFont
import sys


class Ui_ROSA_VETROV(object):
    QToolTip.setFont(QFont('TimesNewRoman', 10))
    
    def __init__(self):
        self.start_calc = None
        self.meteostation = None
        self.centralwidget = None
        self.stations = []  # Список объектов WeatherStation
        self.default_stations_loaded = False  # Флаг загрузки станций по умолчанию
    
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
        
        # Кнопка загрузки CSV файла
        self.load_csv_button = QtWidgets.QPushButton(self.centralwidget)
        self.load_csv_button.setGeometry(QtCore.QRect(10, 10, 261, 41))
        self.load_csv_button.setStyleSheet("""
            background-color: rgb(100, 150, 255);
            color: rgb(255, 255, 255);
            font-weight: bold;
        """)
        self.load_csv_button.setObjectName("load_csv_button")
        self.load_csv_button.setText("Загрузить CSV файл")
        self.load_csv_button.setToolTip('Загрузить CSV файл с данными метеостанции')
        
        # Кнопка удаления выбранной станции
        self.remove_station_button = QtWidgets.QPushButton(self.centralwidget)
        self.remove_station_button.setGeometry(QtCore.QRect(10, 60, 261, 41))
        self.remove_station_button.setStyleSheet("""
            background-color: rgb(255, 100, 100);
            color: rgb(255, 255, 255);
            font-weight: bold;
        """)
        self.remove_station_button.setObjectName("remove_station_button")
        self.remove_station_button.setText("Удалить выбранную станцию")
        self.remove_station_button.setToolTip('Удалить выбранную метеостанцию из списка')
        
        # Выбор метеостанции
        self.meteostation = QtWidgets.QComboBox(self.centralwidget)
        self.meteostation.setGeometry(QtCore.QRect(10, 110, 261, 41))
        self.meteostation.setStyleSheet("""
            background-color: rgb(255, 255, 255);
            color: rgb(0, 0, 0);
        """)
        self.meteostation.setObjectName("meteostation")
        self.meteostation.setToolTip('Выберите метеостанцию из списка загруженных')
        
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
        
        # Метки для дат
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(280, 440, 191, 31))
        self.label.setStyleSheet("""
            background-color: transparent;
            color: rgb(0, 0, 0);
            font-weight: bold;
            font-size: 14px;
        """)
        self.label.setObjectName("label")
        self.label.setText("Выберите начальную дату")
        
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(280, 490, 191, 31))
        self.label_2.setStyleSheet("""
            background-color: transparent;
            color: rgb(0, 0, 0);
            font-weight: bold;
            font-size: 14px;
        """)
        self.label_2.setObjectName("label_2")
        self.label_2.setText("Выберите конечную дату")
        
        # Выбор типа розы ветров
        self.comboBox = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox.setGeometry(QtCore.QRect(10, 170, 261, 41))
        self.comboBox.setStyleSheet("""
                    background-color: rgb(255, 255, 255);
                    color: rgb(0, 0, 0);
                """)
        self.comboBox.setObjectName("comboBox")
        self.comboBox.addItem("Роза ветров цветная")
        self.comboBox.addItem("Роза ветров черно-белая")
        self.comboBox.addItem("Роза ветров контур")
        self.comboBox.setToolTip('Выберите представление розы ветров')
        
        # Область для графика
        self.graphicsView = QtWidgets.QGraphicsView(self.centralwidget)
        self.graphicsView.setGeometry(QtCore.QRect(390, 50, 331, 301))
        self.graphicsView.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.graphicsView.setObjectName("graphicsView")
        self.graphicsView.setToolTip('Здесь будет представлен график розы ветров')
        
        # Начальная дата
        self.n_data = QtWidgets.QDateEdit(self.centralwidget)
        self.n_data.setGeometry(QtCore.QRect(280, 470, 194, 22))
        self.n_data.setStyleSheet("""
            background-color: rgb(255, 255, 255);
            color: rgb(0, 0, 0);
        """)
        self.n_data.setObjectName("n_data")
        self.n_data.setMinimumDate(QtCore.QDate(2006, 9, 1))
        self.n_data.setMaximumDate(QtCore.QDate(2024, 9, 1))
        self.n_data.setToolTip('Начальная дата для расчета')
        
        # Конечная дата
        self.k_data = QtWidgets.QDateEdit(self.centralwidget)
        self.k_data.setGeometry(QtCore.QRect(280, 520, 194, 22))
        self.k_data.setStyleSheet("""
            background-color: rgb(255, 255, 255);
            color: rgb(0, 0, 0);
        """)
        self.k_data.setObjectName("k_data")
        self.k_data.setMinimumDate(QtCore.QDate(2010, 8, 31))
        self.k_data.setMaximumDate(QtCore.QDate(2025, 8, 31))
        self.k_data.setToolTip('Конечная дата для расчета')
        
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
        
        # Чекбоксы для условий
        self.checkBox = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox.setGeometry(QtCore.QRect(300, 330, 81, 21))
        self.checkBox.setObjectName("checkBox")
        self.checkBox.setText("Снег")
        self.checkBox.setToolTip('Учитывать только случаи с осадками в виде снега')
        
        self.checkBox_2 = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox_2.setGeometry(QtCore.QRect(300, 370, 141, 20))
        self.checkBox_2.setObjectName("checkBox_2")
        self.checkBox_2.setText("Ветер ≥3 м/с")
        self.checkBox_2.setToolTip('Учитывать только случаи со скоростью ветра 3 м/с и более')
        
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
            self.load_csv_button, self.remove_station_button, self.meteostation,
            self.start_calc, self.label, self.label_2, self.comboBox,
            self.graphicsView, self.n_data, self.k_data, self.r_cond,
            self.exit_program, self.checkBox, self.checkBox_2, self.rose_name
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
        
        self.statusbar = QtWidgets.QStatusBar(ROSA_VETROV)
        self.statusbar.setObjectName("statusbar")
        ROSA_VETROV.setStatusBar(self.statusbar)
        
        self.retranslateUi(ROSA_VETROV)
        QtCore.QMetaObject.connectSlotsByName(ROSA_VETROV)
        
        # Подключение сигналов
        self.set()
    
    def retranslateUi(self, ROSA_VETROV):
        _translate = QtCore.QCoreApplication.translate
        ROSA_VETROV.setWindowTitle(_translate("ROSA_VETROV", "Программа построения розы ветров"))
        self.rose_name.setText(_translate("ROSA_VETROV", "Выберите метеостанцию и параметры"))
    
    def set(self):
        """Подключение сигналов к слотам"""
        self.start_calc.clicked.connect(self.equal)
        self.exit_program.clicked.connect(app.quit)
        self.load_csv_button.clicked.connect(self.load_csv_file)
        self.remove_station_button.clicked.connect(self.remove_selected_station)
        self.action_load_csv.triggered.connect(self.load_csv_file)
        self.action_exit.triggered.connect(app.quit)
        self.action_about.triggered.connect(self.show_about)
        
        # Загружаем станции по умолчанию при первом запуске
        if not self.default_stations_loaded:
            self.load_default_stations()
            self.default_stations_loaded = True
    
    def load_default_stations(self):
        """Загрузка станций по умолчанию, если файлы существуют"""
        default_files = {
            'Мценск': 'archiv_mcensk_05_24.csv',
            'Орел': 'arсhiv_orel_05_24.csv',
            'Верховье': 'arсhiv_verh_05_21.csv'
        }
        
        for city, filename in default_files.items():
            if os.path.exists(filename):
                try:
                    station = wind_plot.WeatherStation(filename, city)
                    self.stations.append(station)
                    self.meteostation.addItem(city)
                    print(f"Загружена станция по умолчанию: {city}")
                except Exception as e:
                    print(f"Ошибка при загрузке станции {city}: {e}")
    
    def load_csv_file(self):
        """Загрузка CSV файла с данными метеостанции"""
        file_path, _ = QFileDialog.getOpenFileName(
            self.centralwidget,
            "Выберите CSV файл с данными метеостанции",
            "",
            "CSV files (*.csv);;All files (*.*)"
        )
        
        if file_path:
            # Запрос названия города
            city_name, ok = QInputDialog.getText(
                self.centralwidget,
                "Название метеостанции",
                "Введите название города (в именительном падеже):"
            )
            
            if ok and city_name:
                # Запрос названия в предложном падеже
                case_city, ok_case = QInputDialog.getText(
                    self.centralwidget,
                    "Название в предложном падеже",
                    f"Введите название города в предложном падеже (например, 'в {city_name}'):",
                    text=f"{city_name}"
                )
                
                if ok_case:
                    try:
                        # Создание объекта метеостанции
                        station = wind_plot.WeatherStation(file_path, city_name, case_city)
                        self.stations.append(station)
                        self.meteostation.addItem(city_name)
                        
                        # Обновление статуса
                        self.statusbar.showMessage(f"Загружена станция: {city_name}", 3000)
                        QMessageBox.information(
                            self.centralwidget,
                            "Станция загружена",
                            f"Метеостанция '{city_name}' успешно загружена."
                        )
                    except Exception as e:
                        QMessageBox.critical(
                            self.centralwidget,
                            "Ошибка",
                            f"Не удалось загрузить файл: {str(e)}"
                        )
    
    def remove_selected_station(self):
        """Удаление выбранной метеостанции из списка"""
        if not self.stations:
            QMessageBox.warning(self.centralwidget, "Предупреждение", "Список станций пуст.")
            return
        
        selected_index = self.meteostation.currentIndex()
        if selected_index >= 0 and selected_index < len(self.stations):
            station = self.stations[selected_index]
            reply = QMessageBox.question(
                self.centralwidget,
                "Подтверждение",
                f"Вы уверены, что хотите удалить станцию '{station.city_name}'?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.stations.pop(selected_index)
                self.meteostation.removeItem(selected_index)
                self.statusbar.showMessage(f"Станция удалена", 3000)
    
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
        
        # Проверка наличия станций
        if not self.stations:
            self.r_cond.setText("Ошибка: Нет загруженных метеостанций.\n"
                                "Загрузите CSV файл с данными.")
            return
        
        # Получение выбранной станции
        selected_index = self.meteostation.currentIndex()
        if selected_index < 0 or selected_index >= len(self.stations):
            self.r_cond.setText("Ошибка: Выберите метеостанцию из списка.")
            return
        
        station = self.stations[selected_index]
        
        # Обновление названия в интерфейсе
        s = f"{station.city_name}: {self.comboBox.currentText()}"
        self.rose_name.setText(s)
        
        # Получение параметров
        a = self.checkBox.isChecked()  # Снег
        b = self.checkBox_2.isChecked()  # Ветер ≥3 м/с
        
        # Проверка дат
        if self.n_data.date().toString('yyyy.MM.dd') > self.k_data.date().toString('yyyy.MM.dd'):
            self.r_cond.setText('Ошибка: Конечная дата меньше начальной')
            return
        
        # Получение дат
        date_n = self.n_data.date().toString('dd.MM.yyyy')
        date_k = self.k_data.date().toString('dd.MM.yyyy')
        
        # Формирование строки условий
        s_r = 'Расчеты проводятся для: \n' + 'метеостанции: ' + \
              station.city_name + '\n'
        s_r = s_r + 'c ' + date_n + ' по ' + date_k + ' \n'
        
        if a:
            s_1 = 'при наличии осадков в виде снега \n'
        else:
            s_1 = 'независимо от осадков \n'
        
        if b:
            s_1 = s_1 + 'и при ветре 3 и более м/с \n'
        else:
            s_1 = s_1 + 'независимо от скорости ветра \n'
        
        s_r = s_r + s_1
        
        # Словарь для типов розы ветров
        di = {
            "Роза ветров черно-белая": 0,
            "Роза ветров цветная": 1,
            "Роза ветров контур": 2
        }

        # Получение типа розы ветров
        rose_type = di.get(self.comboBox.currentText(), 1)

        # Выполнение расчета
        try:
            file_name = wind_plot.obr_file(
                station,
                date_n, date_k, a, b,
                rose_type
            )
    
            s_r = s_r + 'Роза ветров и таблица сохранены в файле: \n' + file_name
            self.r_cond.setText(s_r)
    
            # Отображение графика
            metadata = station.get_metadata()
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
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)

    # Настройка стиля приложения
    app.setStyle('Fusion')

    # Создание главного окна
    rose_wind = QtWidgets.QMainWindow()
    ui = Ui_ROSA_VETROV()
    ui.setupUi(rose_wind)

    # Установка иконки (если есть)
    try:
        app.setWindowIcon(QtGui.QIcon('wind_rose_icon.png'))
    except:
        pass

    # Показ окна
    rose_wind.show()

    # Запуск приложения
    sys.exit(app.exec_())