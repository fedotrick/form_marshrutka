import sys
import os
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLineEdit,
    QPushButton, QMessageBox, QLabel, QComboBox, QDateEdit,
    QGridLayout, QGroupBox
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
        
        # Initialize all form fields
        self.учетный_номер = QComboBox()
        self.наименование_отливки = QLineEdit()
        self.тип_эксперемента = QLineEdit()
        self.сборка_кластера_дата = QDateEdit()
        self.сборка_кластера_дата.setDate(QDate.currentDate())
        self.сборка_кластера_дата.setCalendarPopup(True)
        self.сборка_кластера_дата.setDisplayFormat("dd.MM.yyyy")
        self.сборка_кластера_специалист = QComboBox()
        self.сборка_кластера_специалист.addItems(["Иванов", "Петров", "Сидоров", "Смирнов"])
        self.сборка_кластера_количество = QLineEdit()
        self.контроль_сборки_кластера_дата_выставления = QDateEdit()
        self.контроль_сборки_кластера_дата_выставления.setDate(QDate.currentDate())
        self.контроль_сборки_кластера_дата_выставления.setCalendarPopup(True)
        self.контроль_сборки_кластера_дата_выставления.setDisplayFormat("dd.MM.yyyy")
        self.контроль_сборки_кластера_время_выставления = QLineEdit()
        self.контроль_сборки_кластера_специалист = QComboBox()
        self.контроль_сборки_кластера_специалист.addItems(["Иванов", "Петров", "Сидоров", "Смирнов"])
        self.болгарка_дата = QDateEdit()
        self.болгарка_дата.setDate(QDate.currentDate())
        self.болгарка_дата.setCalendarPopup(True)
        self.болгарка_дата.setDisplayFormat("dd.MM.yyyy")
        self.болгарка_специалист = QComboBox()
        self.болгарка_специалист.addItems(["Иванов", "Петров", "Сидоров", "Смирнов"])
        self.термообработка_специалист = QComboBox()
        self.термообработка_специалист.addItems(["Иванов", "Петров", "Сидоров", "Смирнов"])
        self.дробеметная_обработка_специалист = QComboBox()
        self.дробеметная_обработка_специалист.addItems(["Иванов", "Петров", "Сидоров", "Смирнов"])
        self.зачистка_корона_специалист = QComboBox()
        self.зачистка_корона_специалист.addItems(["Иванов", "Петров", "Сидоров", "Смирнов"])
        self.зачистка_лапа_специалист = QComboBox()
        self.зачистка_лапа_специалист.addItems(["Иванов", "Петров", "Сидоров", "Смирнов"])
        self.зачистка_питатель_специалист = QComboBox()
        self.зачистка_питатель_специалист.addItems(["Иванов", "Петров", "Сидоров", "Смирнов"])
        self.примечание = QLineEdit()
        
        # Apply stylesheet
        self.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                font-family: 'Segoe UI', Arial;
                font-size: 11px;
            }
            QLabel {
                color: #333333;
                padding: 3px;
            }
            QLineEdit, QComboBox {
                padding: 5px;
                border: 1px solid #cccccc;
                border-radius: 3px;
                background-color: #ffffff;
                min-width: 180px;
                margin-bottom: 5px;
            }
            QDateEdit {
                padding: 5px 25px 5px 5px;
                border: 1px solid #cccccc;
                border-radius: 3px;
                background-color: #ffffff;
                min-width: 180px;
                margin-bottom: 5px;
                font-size: 12px;
            }
            QDateEdit::drop-down {
                border: none;
                width: 20px;
                background-color: transparent;
            }
            QDateEdit::down-arrow {
                image: url(data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxNiIgaGVpZ2h0PSIxNiIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9IiM0YTkwZTIiIHN0cm9rZS13aWR0aD0iMiIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIj48cmVjdCB4PSIzIiB5PSI0IiB3aWR0aD0iMTgiIGhlaWdodD0iMTgiIHJ4PSIyIiByeT0iMiI+PC9yZWN0PjxsaW5lIHgxPSIxNiIgeTE9IjIiIHgyPSIxNiIgeTI9IjYiPjwvbGluZT48bGluZSB4MT0iOCIgeTE9IjIiIHgyPSI4IiB5Mj0iNiI+PC9saW5lPjxsaW5lIHgxPSIzIiB5MT0iMTAiIHgyPSIyMSIgeTI9IjEwIj48L2xpbmU+PC9zdmc+);
                width: 16px;
                height: 16px;
            }
            QDateEdit:focus, QLineEdit:focus, QComboBox:focus {
                border: 2px solid #4a90e2;
                outline: none;
            }
            QDateEdit:hover, QLineEdit:hover, QComboBox:hover {
                border: 1px solid #4a90e2;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QPushButton {
                background-color: #4a90e2;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 4px;
                font-weight: bold;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #357abd;
            }
            QPushButton:pressed {
                background-color: #2d6da3;
            }
            QGroupBox {
                font-weight: bold;
                background-color: #f8f9fa;
                border-radius: 4px;
                margin-top: 10px;
                padding-top: 10px;
            }
        """)

        # Create main layout
        main_layout = QGridLayout()
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(8)

        # Header
        header = QWidget()
        header.setStyleSheet("background-color: #4a90e2; border-radius: 4px; margin-bottom: 10px;")
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(5, 5, 5, 5)
        title = QLabel("МАРШРУТНАЯ КАРТА")
        title.setStyleSheet("font-size: 18px; font-weight: bold; color: white; padding: 8px;")
        title.setAlignment(Qt.AlignCenter)
        header_layout.addWidget(title)
        main_layout.addWidget(header, 0, 0, 1, 4)

        # Create and populate group boxes
        # Основная информация
        info_group = QGroupBox("Основная информация")
        info_layout = QVBoxLayout(info_group)
        info_layout.addWidget(QLabel("Учетный номер:"))
        info_layout.addWidget(self.учетный_номер)
        info_layout.addWidget(QLabel("Наименование отливки:"))
        info_layout.addWidget(self.наименование_отливки)
        info_layout.addWidget(QLabel("Тип эксперимента:"))
        info_layout.addWidget(self.тип_эксперемента)
        main_layout.addWidget(info_group, 1, 0)

        # Сборка кластера
        assembly_group = QGroupBox("Сборка кластера")
        assembly_layout = QVBoxLayout(assembly_group)
        assembly_layout.addWidget(QLabel("Дата сборки:"))
        assembly_layout.addWidget(self.сборка_кластера_дата)
        assembly_layout.addWidget(QLabel("Специалист:"))
        assembly_layout.addWidget(self.сборка_кластера_специалист)
        assembly_layout.addWidget(QLabel("Количество:"))
        assembly_layout.addWidget(self.сборка_кластера_количество)
        main_layout.addWidget(assembly_group, 1, 1)

        # Контроль сборки
        control_group = QGroupBox("Контроль сборки")
        control_layout = QVBoxLayout(control_group)
        control_layout.addWidget(QLabel("Дата выставления:"))
        control_layout.addWidget(self.контроль_сборки_кластера_дата_выставления)
        control_layout.addWidget(QLabel("Время выставления:"))
        control_layout.addWidget(self.контроль_сборки_кластера_время_выставления)
        control_layout.addWidget(QLabel("Специалист контроля:"))
        control_layout.addWidget(self.контроль_сборки_кластера_специалист)
        main_layout.addWidget(control_group, 1, 2)

        # Обработка
        processing_group = QGroupBox("Обработка")
        processing_layout = QVBoxLayout(processing_group)
        processing_layout.addWidget(QLabel("Дата болгарки:"))
        processing_layout.addWidget(self.болгарка_дата)
        processing_layout.addWidget(QLabel("Специалист болгарки:"))
        processing_layout.addWidget(self.болгарка_специалист)
        processing_layout.addWidget(QLabel("Специалист термообработки:"))
        processing_layout.addWidget(self.термообработка_специалист)
        processing_layout.addWidget(QLabel("Специалист дробеметной обработки:"))
        processing_layout.addWidget(self.дробеметная_обработка_специалист)
        main_layout.addWidget(processing_group, 1, 3)

        # Зачистка
        cleaning_group = QGroupBox("Зачистка")
        cleaning_layout = QVBoxLayout(cleaning_group)
        cleaning_layout.addWidget(QLabel("Специалист зачистки короны:"))
        cleaning_layout.addWidget(self.зачистка_корона_специалист)
        cleaning_layout.addWidget(QLabel("Специалист зачистки лапы:"))
        cleaning_layout.addWidget(self.зачистка_лапа_специалист)
        cleaning_layout.addWidget(QLabel("Специалист зачистки питателя:"))
        cleaning_layout.addWidget(self.зачистка_питатель_специалист)
        main_layout.addWidget(cleaning_group, 2, 0, 1, 2)

        # Примечание
        note_group = QGroupBox("Дополнительно")
        note_layout = QVBoxLayout(note_group)
        note_layout.addWidget(QLabel("Примечание:"))
        note_layout.addWidget(self.примечание)
        main_layout.addWidget(note_group, 2, 2, 1, 2)

        # Save button
        self.save_button = QPushButton("Сохранить", self)
        self.save_button.setStyleSheet("""
            background-color: #4CAF50;
            color: white;
            padding: 12px 24px;
            border-radius: 4px;
            font-size: 16px;
            font-weight: bold;
            margin-top: 10px;
        """)
        self.save_button.clicked.connect(self.save_data)
        main_layout.addWidget(self.save_button, 3, 0, 1, 4)

        self.setLayout(main_layout)

        # Load initial data
        current_account_numbers = load_account_numbers('plavka.xlsx')
        self.учетный_номер.addItems(current_account_numbers)
        self.учетный_номер.currentTextChanged.connect(self.update_experiment_details)

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
        зачищка_лапа_специалист = self.зачистка_лапа_специалист.currentText().strip()
        зачищка_питатель_специалист = self.зачистка_питатель_специалист.currentText().strip()
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
        self.болгарка_дата.setDate(QDate.currentDate())
        self.болгарка_специалист.setCurrentIndex(-1)  # Сброс выбора
        self.термообработка_специалист.setCurrentIndex(-1)  # Сброс выбора
        self.дробеметная_обработка_специалист.setCurrentIndex(-1)  # Сброс выбора
        self.зачистка_корона_специалист.setCurrentIndex(-1)  # Сброс выбора
        self.зачистка_лапа_специалист.setCurrentIndex(-1)  # Сброс выбора
        self.зачистка_питатель_специалист.setCurrentIndex(-1)  # Сброс выбора
        self.наименование_отливки.clear()
        self.тип_эксперемента.clear()
        self.примечание.clear()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
