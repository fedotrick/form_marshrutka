import sys
import os
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLineEdit,
    QPushButton, QMessageBox, QLabel, QComboBox, QDateEdit,
    QGridLayout, QScrollArea
)
from PySide6 import QtGui
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QFont
from openpyxl import Workbook, load_workbook

# Функция для загрузки учетных номеров из Excel
def load_account_numbers(file_name):
    workbook = load_workbook(file_name)
    sheet = workbook.active
    account_numbers = []
    
    # Получаем уже использованные номера из marshrutka.xlsx
    used_numbers = set()
    if os.path.exists('marshrutka.xlsx'):
        marshrutka_wb = load_workbook('marshrutka.xlsx')
        if "Records" in marshrutka_wb.sheetnames:
            marshrutka_sheet = marshrutka_wb["Records"]
            for row in marshrutka_sheet.iter_rows(min_row=2, values_only=True):
                if row[6]:  # Учетный номер в 7-м столбце (индекс 6)
                    used_numbers.add(row[6])
        marshrutka_wb.close()
    
    # Фильтруем номера из plavka.xlsx
    for row in sheet.iter_rows(min_row=2, values_only=True):
        account_number = row[1]  # Учетный номер во втором столбце
        if (account_number 
            and "/25" in str(account_number)  # Содержит "/25"
            and account_number not in used_numbers):  # Отсутствует в marshrutka.xlsx
            account_numbers.append(account_number)
    
    return sorted(account_numbers)  # Возвращаем отсортированный список

# Функция для сохранения данных в Excel
def save_to_excel(сборка_кластера_дата, сборка_кластера_специалист, сборка_кластера_количество,
                  контроль_сборки_кластера_дата_выставления, контроль_сборки_кластера_время_выставления,
                  контроль_сборки_кластера_специалист, учетный_номер, наименование_отливки, тип_эксперемента,
                  болгарка_дата, болгарка_специалист, термообработка_специалист,
                  дробеметная_обработка_специалист, зачищка_корона_специалист,
                  зачищка_лапа_специалист, зачищка_питатель_специалист, примечание):
    try:
        if os.path.exists('marshrutka.xlsx'):
            wb = load_workbook('marshrutka.xlsx')
            # Получаем лист Records или создаем его, если не существует
            if "Records" not in wb.sheetnames:
                ws = wb.create_sheet("Records")
                headers = ['Дата сборки', 'Специалист сборки', 'Количество', 
                          'Дата выставления', 'Время выставления', 'Специалист контроля',
                          'Учетный номер', 'Наименование отливки', 'Тип эксперимента',
                          'Дата болгарки', 'Специалист болгарки', 'Специалист термообработки',
                          'Специалист дробеметной обработки', 'Специалист зачистки короны',
                          'Специалист зачистки лапы', 'Специалист зачистки питателя', 'Примечание']
                for col, header in enumerate(headers, start=1):
                    ws.cell(row=1, column=col, value=header)
            else:
                ws = wb["Records"]
            
            next_row = ws.max_row + 1
            
            data = [сборка_кластера_дата, сборка_кластера_специалист, сборка_кластера_количество,
                    контроль_сборки_кластера_дата_выставления, контроль_сборки_кластера_время_выставления,
                    контроль_сборки_кластера_специалист, учетный_номер, наименование_отливки, тип_эксперемента,
                    болгарка_дата, болгарка_специалист, термообработка_специалист,
                    дробеметная_обработка_специалист, зачищка_корона_специалист,
                    зачищка_лапа_специалист, зачищка_питатель_специалист, примечание]
            
            for col, value in enumerate(data, start=1):
                cell = ws.cell(row=next_row, column=col)
                cell.value = value
                
                if col in [1, 4, 10]:  # Колонки с датами
                    cell.number_format = 'DD.MM.YYYY'
                elif col == 5:  # Колонка с временем
                    cell.number_format = 'HH:MM'
            
            wb.save('marshrutka.xlsx')
            wb.close()
        else:
            # Создаем новый файл если он не существует
            wb = Workbook()
            ws = wb.active
            ws.title = "Records"
            
            # Добавляем заголовки
            headers = ['Дата сборки', 'Специалист сборки', 'Количество', 
                      'Дата выставления', 'Время выставления', 'Специалист контроля',
                      'Учетный номер', 'Наименование отливки', 'Тип эксперимента',
                      'Дата болгарки', 'Специалист болгарки', 'Специалист термообработки',
                      'Специалист дробеметной обработки', 'Специалист зачистки короны',
                      'Специалист зачистки лапы', 'Специалист зачистки питателя', 'Примечание']
            for col, header in enumerate(headers, start=1):
                ws.cell(row=1, column=col, value=header)
            
            # Добавляем данные
            data = [сборка_кластера_дата, сборка_кластера_специалист, сборка_кластера_количество,
                    контроль_сборки_кластера_дата_выставления, контроль_сборки_кластера_время_выставления,
                    контроль_сборки_кластера_специалист, учетный_номер, наименование_отливки, тип_эксперемента,
                    болгарка_дата, болгарка_специалист, термообработка_специалист,
                    дробеметная_обработка_специалист, зачищка_корона_специалист,
                    зачищка_лапа_специалист, зачищка_питатель_специалист, примечание]
            
            for col, value in enumerate(data, start=1):
                cell = ws.cell(row=2, column=col)
                cell.value = value
                
                if col in [1, 4, 10]:  # Колонки с датами
                    cell.number_format = 'DD.MM.YYYY'
                elif col == 5:  # Колонка с временем
                    cell.number_format = 'HH:MM'
            
            wb.save('marshrutka.xlsx')
            wb.close()
    except Exception as e:
        raise Exception(f"Ошибка при сохранении в Excel: {str(e)}")


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Электронная маршрутная карта")
        # Установка фиксированного размера окна
        self.setFixedSize(800, 900)
        
        # Основные цвета Aura
        self.BRAND_COLOR = "#0176D3"
        self.TEXT_COLOR = "#181818"
        self.BORDER_COLOR = "#DDDBDA"
        self.BG_COLOR = "#F3F3F3"
        self.SUCCESS_COLOR = "#45C65A"
        
        # Установка основного стиля окна
        self.setStyleSheet(f"""
            QWidget {{
                background-color: {self.BG_COLOR};
                font-family: 'Segoe UI', Arial;
                font-size: 13px;
                color: {self.TEXT_COLOR};
            }}
            QLabel {{
                padding: 4px 0;
                font-size: 13px;
            }}
            QLineEdit, QComboBox, QDateEdit {{
                padding: 4px 8px;
                border: 1px solid {self.BORDER_COLOR};
                border-radius: 4px;
                background-color: white;
                min-height: 24px;
            }}
            QLineEdit:focus, QComboBox:focus, QDateEdit:focus {{
                border: 2px solid {self.BRAND_COLOR};
                outline: none;
            }}
            QComboBox::drop-down {{
                border: none;
                width: 24px;
            }}
            QPushButton {{
                background-color: {self.BRAND_COLOR};
                color: white;
                padding: 6px 16px;
                border: none;
                border-radius: 4px;
                font-weight: bold;
                min-height: 24px;
            }}
            QPushButton:hover {{
                background-color: #014486;
            }}
            QPushButton:pressed {{
                background-color: #032D60;
            }}
        """)

        # Списки специалистов
        scleyks = ["Буцик", "Минакова", "Ротарь", "Чернова", "Чупахина"]
        controlers = ["Елхова", "Шестункина", "Романцева"]
        bolgar = ["Ахмаджонов", "Отаназаров", "Косимов", "Туичев",
                 "Машрапов", "Эргашев", "Самиев", "Исмаилов"]
        termob = ["Эгамов", "Аюбов"]
        drobem = ["Эгамов", "Аюбов"]
        zachistka = ["Буцик", "Минакова", "Ротарь", "Чернова", "Чупахина"]

        # Инициализация всех полей формы
        self.сборка_кластера_дата = QDateEdit(self)
        self.сборка_кластера_дата.setDisplayFormat("dd.MM.yyyy")
        self.сборка_кластера_дата.setCalendarPopup(True)
        self.сборка_кластера_дата.setDate(QDate.currentDate())

        self.сборка_кластера_специалист = QComboBox(self)
        self.сборка_кластера_специалист.addItems(scleyks)
        self.сборка_кластера_специалист.setCurrentIndex(-1)
        self.сборка_кластера_специалист.setPlaceholderText("Специалист по сборке кластера")

        self.сборка_кластера_количество = QLineEdit(self)
        self.сборка_кластера_количество.setPlaceholderText("Количество кластера")

        self.контроль_сборки_кластера_дата_выставления = QDateEdit(self)
        self.контроль_сборки_кластера_дата_выставления.setDisplayFormat("dd.MM.yyyy")
        self.контроль_сборки_кластера_дата_выставления.setCalendarPopup(True)
        self.контроль_сборки_кластера_дата_выставления.setDate(QDate.currentDate())

        self.контроль_сборки_кластера_время_выставления = QLineEdit(self)
        self.контроль_сборки_кластера_время_выставления.setPlaceholderText("Время (ЧЧ:ММ)")
        self.контроль_сборки_кластера_время_выставления.setInputMask("99:99")

        self.контроль_сборки_кластера_специалист = QComboBox(self)
        self.контроль_сборки_кластера_специалист.addItems(controlers)
        self.контроль_сборки_кластера_специалист.setCurrentIndex(-1)
        self.контроль_сборки_кластера_специалист.setPlaceholderText("Специалист по контролю кластера")

        self.учетный_номер = QComboBox(self)
        self.учетный_номер.addItems(load_account_numbers('plavka.xlsx'))
        self.учетный_номер.currentTextChanged.connect(self.update_experiment_details)

        self.наименование_отливки = QLineEdit(self)
        self.наименование_отливки.setPlaceholderText("Наименование отливки")

        self.тип_эксперемента = QLineEdit(self)
        self.тип_эксперемента.setPlaceholderText("Тип эксперимента")

        self.болгарка_дата = QDateEdit(self)
        self.болгарка_дата.setDisplayFormat("dd.MM.yyyy")
        self.болгарка_дата.setCalendarPopup(True)
        self.болгарка_дата.setDate(QDate.currentDate())

        self.болгарка_специалист = QComboBox(self)
        self.болгарка_специалист.addItems(bolgar)
        self.болгарка_специалист.setCurrentIndex(-1)
        self.болгарка_специалист.setPlaceholderText("Специалист по резке")

        self.термообработка_специалист = QComboBox(self)
        self.термообработка_специалист.addItems(termob)
        self.термообработка_специалист.setCurrentIndex(-1)
        self.термообработка_специалист.setPlaceholderText("Специалист по термообработке")

        self.дробеметная_обработка_специалист = QComboBox(self)
        self.дробеметная_обработка_специалист.addItems(drobem)
        self.дробеметная_обработка_специалист.setCurrentIndex(-1)
        self.дробеметная_обработка_специалист.setPlaceholderText("Специалист по дробеметной обработке")

        self.зачистка_корона_специалист = QComboBox(self)
        self.зачистка_корона_специалист.addItems(zachistka)
        self.зачистка_корона_специалист.setCurrentIndex(-1)
        self.зачистка_корона_специалист.setPlaceholderText("Специалист по зачистке короны")

        self.зачистка_лапа_специалист = QComboBox(self)
        self.зачистка_лапа_специалист.addItems(zachistka)
        self.зачистка_лапа_специалист.setCurrentIndex(-1)
        self.зачистка_лапа_специалист.setPlaceholderText("Специалист по зачистке лапы")

        self.зачистка_питатель_специалист = QComboBox(self)
        self.зачистка_питатель_специалист.addItems(zachistka)
        self.зачистка_питатель_специалист.setCurrentIndex(-1)
        self.зачистка_питатель_специалист.setPlaceholderText("Специалист по зачистке питателя")

        self.примечание = QLineEdit(self)
        self.примечание.setPlaceholderText("Примечание")

        # Создаем основной layout
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(8)
        main_layout.setContentsMargins(16, 16, 16, 16)

        # Заголовок
        title = QLabel("МАРШРУТНАЯ КАРТА")
        title.setStyleSheet(f"""
            font-size: 20px;
            font-weight: bold;
            color: {self.BRAND_COLOR};
            padding: 0 0 8px 0;
            border-bottom: 2px solid {self.BORDER_COLOR};
            margin-bottom: 8px;
        """)
        main_layout.addWidget(title)

        # Создаем область прокрутки
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet(f"QScrollArea {{ border: none; background-color: {self.BG_COLOR}; }}")
        
        # Создаем виджет для содержимого прокрутки
        scroll_content = QWidget()
        grid_layout = QGridLayout(scroll_content)
        grid_layout.setSpacing(8)
        
        current_row = 0

        # Размещаем элементы в сетке
        def add_field(label_text, widget, row, col=0):
            label = QLabel(label_text)
            grid_layout.addWidget(label, row, col)
            grid_layout.addWidget(widget, row, col + 1)

        # Добавляем поля в сетку
        add_field("Дата сборки:", self.сборка_кластера_дата, current_row); current_row += 1
        add_field("Специалист сборки:", self.сборка_кластера_специалист, current_row); current_row += 1
        add_field("Количество:", self.сборка_кластера_количество, current_row); current_row += 1
        add_field("Дата выставления:", self.контроль_сборки_кластера_дата_выставления, current_row); current_row += 1
        add_field("Время выставления:", self.контроль_сборки_кластера_время_выставления, current_row); current_row += 1
        add_field("Специалист контроля:", self.контроль_сборки_кластера_специалист, current_row); current_row += 1
        add_field("Учетный номер:", self.учетный_номер, current_row); current_row += 1
        add_field("Наименование отливки:", self.наименование_отливки, current_row); current_row += 1
        add_field("Тип эксперимента:", self.тип_эксперемента, current_row); current_row += 1
        add_field("Дата болгарки:", self.болгарка_дата, current_row); current_row += 1
        add_field("Специалист болгарки:", self.болгарка_специалист, current_row); current_row += 1
        add_field("Специалист термообработки:", self.термообработка_специалист, current_row); current_row += 1
        add_field("Специалист дробеметной обработки:", self.дробеметная_обработка_специалист, current_row); current_row += 1
        add_field("Специалист зачистки короны:", self.зачистка_корона_специалист, current_row); current_row += 1
        add_field("Специалист зачистки лапы:", self.зачистка_лапа_специалист, current_row); current_row += 1
        add_field("Специалист зачистки питателя:", self.зачистка_питатель_специалист, current_row); current_row += 1
        
        # Добавляем поле примечания
        grid_layout.addWidget(QLabel("Примечание:"), current_row, 0)
        self.примечание.setFixedHeight(60)
        grid_layout.addWidget(self.примечание, current_row, 1)
        
        # Устанавливаем содержимое области прокрутки
        scroll_area.setWidget(scroll_content)
        main_layout.addWidget(scroll_area)

        # Кнопка сохранения
        self.save_button = QPushButton("Сохранить", self)
        self.save_button.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.SUCCESS_COLOR};
                color: white;
                padding: 8px 24px;
                border-radius: 4px;
                font-weight: bold;
                font-size: 14px;
                min-height: 32px;
            }}
            QPushButton:hover {{
                background-color: #2E844A;
            }}
            QPushButton:pressed {{
                background-color: #194E31;
            }}
        """)
        self.save_button.clicked.connect(self.save_data)
        main_layout.addWidget(self.save_button)

        self.setLayout(main_layout)

    def update_experiment_details(self):
        """Обновляет поля на основании выбранного учетного номера"""
        selected_account = self.учетный_номер.currentText()
        if selected_account:
            # Загрузка данных для выбранного учетного номера
            workbook = load_workbook('plavka.xlsx')
            sheet = workbook.active
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if row[1] == selected_account:  # Учетный номер во втором столбце
                    self.наименование_отливки.setText(row[10])  # "Наименование_отливки" в 11-м столбце
                    self.тип_эксперемента.setText(row[11])  # "Тип_эксперемента" в 12-м столбце
                    break

    def validate_time(self, time_str):
        """Проверка корректности ввода времени в формате ЧЧ:ММ"""
        try:
            hours, minutes = map(int, time_str.split(':'))
            if 0 <= hours < 24 and 0 <= minutes < 60:
                return True
        except ValueError:
            return False
        return False

    def save_data(self):
        сборка_кластера_дата = self.сборка_кластера_дата.date().toString("dd.MM.yyyy")
        сборка_кластера_специалист = self.сборка_кластера_специалист.currentText().strip()
        сборка_кластера_количество = self.сборка_кластера_количество.text().strip()
        контроль_сборки_кластера_дата_выставления = self.контроль_сборки_кластера_дата_выставления.date().toString("dd.MM.yyyy")
        контроль_сборки_кластера_время_выставления = self.контроль_сборки_кластера_время_выставления.text().strip()
        контроль_сборки_кластера_специалист = self.контроль_сборки_кластера_специалист.currentText().strip()
        учетный_номер = self.учетный_номер.currentText().strip()
        болгарка_дата = self.болгарка_дата.date().toString("dd.MM.yyyy")
        болгарка_специалист = self.болгарка_специалист.currentText().strip()
        термообработка_специалист = self.термообработка_специалист.currentText().strip()
        дробеметная_обработка_специалист = self.дробеметная_обработка_специалист.currentText().strip()
        зачищка_корона_специалист = self.зачистка_корона_специалист.currentText().strip()
        зачищка_лапа_специалист = self.зачищка_лапа_специалист.currentText().strip()
        зачищка_питатель_специалист = self.зачищка_питатель_специалист.currentText().strip()
        примечание = self.примечание.text().strip()

        # Валидация времени
        if not self.validate_time(контроль_сборки_кластера_время_выставления):
            QMessageBox.warning(self, "Ошибка", "Некорректный ввод времени. Используйте формат ЧЧ:ММ.")
            return

        # Получение наименования отливки и типа эксперимента
        наименование_отливки = self.наименование_отливки.text().strip()
        тип_эксперемента = self.тип_эксперемента.text().strip()

        save_to_excel(сборка_кластера_дата, сборка_кластера_специалист, сборка_кластера_количество,
                       контроль_сборки_кластера_дата_выставления, контроль_сборки_кластера_время_выставления,
                       контроль_сборки_кластера_специалист, учетный_номер, наименование_отливки, тип_эксперемента,
                       болгарка_дата, болгарка_специалист, термообработка_специалист,
                       дробеметная_обработка_специалист, зачищка_корона_специалист,
                       зачищка_лапа_специалист, зачищка_питатель_специалист, примечание)

        QMessageBox.information(self, "Успех", "Данные сохранены в Excel!")

        # Обновляем список учетных номеров
        current_account_numbers = load_account_numbers('plavka.xlsx')
        self.учетный_номер.clear()
        self.учетный_номер.addItems(current_account_numbers)

        # Очистка полей ввода
        self.clear_fields()

    def clear_fields(self):
        self.учетный_номер.setCurrentIndex(0)
        self.сборка_кластера_дата.setDate(QDate.currentDate())
        self.сборка_кластера_специалист.setCurrentIndex(-1)  # Сброс выбора
        self.сборка_кластера_количество.clear()
        self.контроль_сборки_кластера_дата_выставления.setDate(QDate.currentDate())
        self.контроль_сборки_кластера_время_выставления.clear()
        self.контроль_сборки_кластера_специалист.setCurrentIndex(-1)  # Сброс выбора
        self.наименование_отливки.clear()
        self.тип_эксперемента.clear()
        self.болгарка_дата.setDate(QDate.currentDate())
        self.болгарка_специалист.setCurrentIndex(-1)  # Сброс выбора
        self.термообработка_специалист.setCurrentIndex(-1)  # Сброс выбора
        self.дробеметная_обработка_специалист.setCurrentIndex(-1)  # Сброс выбора
        self.зачистка_корона_специалист.setCurrentIndex(-1)  # Сброс выбора
        self.зачистка_лапа_специалист.setCurrentIndex(-1)  # Сброс выбора
        self.зачистка_питатель_специалист.setCurrentIndex(-1)  # Сброс выбора
        self.примечание.clear()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
