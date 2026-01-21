# -*- coding: utf-8 -*-
import sys

from PyQt5 import QtGui, QtWidgets

import ui
import weather_station


def main():
    default_stations = weather_station.load_default_stations(
        [
            weather_station.LoadMeteostationQuery(
                **{
                    "name": "Мценск",
                    "path": "archiv_mcensk_05_24.csv",
                }
            ),
            weather_station.LoadMeteostationQuery(
                **{
                    "name": "Орел",
                    "path": "arсhiv_orel_05_24.csv",
                }
            ),
            weather_station.LoadMeteostationQuery(
                **{
                    "name": "Верховье",
                    "path": "arсhiv_verh_05_21.csv",
                }
            ),
        ]
    )

    app = QtWidgets.QApplication(sys.argv)

    # Настройка стиля приложения
    app.setStyle("Fusion")

    # Создание главного окна
    main_window = QtWidgets.QMainWindow()

    gui = ui.Ui_ROSA_VETROV(default_stations)
    gui.setupUi(main_window)

    # Установка иконки (если есть)
    try:
        app.setWindowIcon(QtGui.QIcon("wind_rose_icon.png"))
    except BaseException:
        pass

    # Показ окна
    main_window.show()

    # Запуск приложения
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
