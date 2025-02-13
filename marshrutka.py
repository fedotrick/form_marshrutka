import sys
import os
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLineEdit,
    QPushButton, QMessageBox, QLabel, QComboBox, QDateEdit
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
        self.setStyleSheet("background-color: #f0f0f0; font-family: Arial; 24px;")

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignTop)

        title = QLabel("МАРШРУТНАЯ КАРТА")
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: black; margin-bottom: 20px;")
        layout.addWidget(title)
        
        # Список склейщиков отливок
        scleyks = [
            "Буцик", "Минакова", "Ротарь", 
            "Чернова", "Чупахина"
        ]
        # Список контролеров сборки кластера
        controlers = [
            "Елхова", "Шестункина", "Романцева"
        ]
        # Список болгарщиков
        bolgar = [
            "Ахмаджонов", "Отаназаров", "Косимов", "Туичев",
            "Машрапов", "Эргашев", "Самиев", "Исмаилов"
        ]
        # Список тербообработчиков
        termob = [
            "Эгамов", "Аюбов"
        ]
        # Список дробеметных обработчиков
        drobem = [
            "Эгамов", "Аюбов"
        ]
        # Список специалистов по зачистки
        zachistka = [
            "Абдуллаев", "Бурхонов", "Матесаев", "Мещерякова",
            "Самиев", "Леонтьева"
        ]
                        
        # Сортировка списков
        scleyks.sort()
        controlers.sort()
        bolgar.sort()
        termob.sort()
        drobem.sort()
        zachistka.sort()

        # Загрузка учетных номеров
        self.учетный_номер = QComboBox(self)
        self.учетный_номер.addItems(load_account_numbers('plavka.xlsx'))
        self.учетный_номер.currentTextChanged.connect(self.update_experiment_details)
        layout.addWidget(self.учетный_номер)

        self.наименование_отливки = QLineEdit(self)
        self.наименование_отливки.setPlaceholderText("Наименование отливки")
        layout.addWidget(self.наименование_отливки)

        self.тип_эксперемента = QLineEdit(self)
        self.тип_эксперемента.setPlaceholderText("Тип эксперимента")
        layout.addWidget(self.тип_эксперемента)
        
        self.сборка_кластера_дата = QDateEdit(self)
        self.сборка_кластера_дата.setDisplayFormat("dd.MM.yyyy")
        self.сборка_кластера_дата.setCalendarPopup(True)
        self.сборка_кластера_дата.setDate(QDate.currentDate())
        layout.addWidget(QLabel("Дата сборки кластера"))
        layout.addWidget(self.сборка_кластера_дата)

        # Заменяем QLineEdit на QComboBox для специалиста сборки кластера
        self.сборка_кластера_специалист = QComboBox(self)
        self.сборка_кластера_специалист.addItems(scleyks)  # Добавляем участников в комбобокс
        self.сборка_кластера_специалист.setFont(QtGui.QFont("Aptos", 10, QtGui.QFont.Bold))
        self.сборка_кластера_специалист.setStyleSheet("color: black;")
        self.сборка_кластера_специалист.setCurrentIndex(-1)  # Ничего не выбрано по умолчанию
        self.сборка_кластера_специалист.setPlaceholderText("Специалист по сборке кластера")
        layout.addWidget(self.сборка_кластера_специалист)

        self.сборка_кластера_количество = QLineEdit(self)
        self.сборка_кластера_количество.setPlaceholderText("Количество кластера")
        layout.addWidget(self.сборка_кластера_количество)

        self.контроль_сборки_кластера_дата_выставления = QDateEdit(self)
        self.контроль_сборки_кластера_дата_выставления.setDisplayFormat("dd.MM.yyyy")
        self.контроль_сборки_кластера_дата_выставления.setCalendarPopup(True)
        self.контроль_сборки_кластера_дата_выставления.setDate(QDate.currentDate())
        layout.addWidget(QLabel("Дата выставления кластера"))
        layout.addWidget(self.контроль_сборки_кластера_дата_выставления)

        self.контроль_сборки_кластера_время_выставления = QLineEdit(self)
        self.контроль_сборки_кластера_время_выставления.setPlaceholderText("Время (ЧЧ:ММ)")
        self.контроль_сборки_кластера_время_выставления.setStyleSheet("padding: 10px; margin-bottom: 10px;")
        self.контроль_сборки_кластера_время_выставления.setInputMask("99:99")
        #self.контроль_сборки_кластера_время_выставления.setMaxLength(5)
        layout.addWidget(QLabel("Время выставления кластера (ЧЧ:ММ):"))
        layout.addWidget(self.контроль_сборки_кластера_время_выставления)

        # Заменяем QLineEdit на QComboBox для специалиста контроля сборки кластера
        self.контроль_сборки_кластера_специалист = QComboBox(self)
        self.контроль_сборки_кластера_специалист.addItems(controlers)  # Добавляем участников в комбобокс
        self.контроль_сборки_кластера_специалист.setFont(QtGui.QFont("Aptos", 10, QtGui.QFont.Bold))
        self.контроль_сборки_кластера_специалист.setStyleSheet("color: black;")
        self.контроль_сборки_кластера_специалист.setCurrentIndex(-1)  # Ничего не выбрано по умолчанию
        self.контроль_сборки_кластера_специалист.setPlaceholderText("Специалист по контролю кластера")
        layout.addWidget(self.контроль_сборки_кластера_специалист)

        self.номер_кластера = QLineEdit(self)
        self.номер_кластера.setPlaceholderText("Номер кластера")
        layout.addWidget(self.номер_кластера)

        self.количество_кластеров = QLineEdit(self)
        self.количество_кластеров.setPlaceholderText("Количество кластеров")
        layout.addWidget(self.количество_кластеров)

        self.болгарка_дата = QDateEdit(self)
        self.болгарка_дата.setDisplayFormat("dd.MM.yyyy")
        self.болгарка_дата.setCalendarPopup(True)
        self.болгарка_дата.setDate(QDate.currentDate())
        layout.addWidget(self.болгарка_дата)

        # Заменяем QLineEdit на QComboBox для специалиста по болгарке
        self.болгарка_специалист = QComboBox(self)
        self.болгарка_специалист.addItems(bolgar)  # Добавляем участников в комбобокс
        self.болгарка_специалист.setFont(QtGui.QFont("Aptos", 10, QtGui.QFont.Bold))
        self.болгарка_специалист.setStyleSheet("color: black;")
        self.болгарка_специалист.setCurrentIndex(-1)  # Ничего не выбрано по умолчанию
        self.болгарка_специалист.setPlaceholderText("Специалист по резке")
        layout.addWidget(self.болгарка_специалист)

        # Заменяем QLineEdit на QComboBox для специалиста по термообработки
        self.термообработка_специалист = QComboBox(self)
        self.термообработка_специалист.addItems(termob)  # Добавляем участников в комбобокс
        self.термообработка_специалист.setFont(QtGui.QFont("Aptos", 10, QtGui.QFont.Bold))
        self.термообработка_специалист.setStyleSheet("color: black;")
        self.термообработка_специалист.setCurrentIndex(-1)  # Ничего не выбрано по умолчанию
        self.термообработка_специалист.setPlaceholderText("Специалист по термообработке")
        layout.addWidget(self.термообработка_специалист)
        
        # Заменяем QLineEdit на QComboBox для специалиста по дробеметной обработке
        self.дробеметная_обработка_специалист = QComboBox(self)
        self.дробеметная_обработка_специалист.addItems(drobem)  # Добавляем участников в комбобокс
        self.дробеметная_обработка_специалист.setFont(QtGui.QFont("Aptos", 10, QtGui.QFont.Bold))
        self.дробеметная_обработка_специалист.setStyleSheet("color: black;")
        self.дробеметная_обработка_специалист.setCurrentIndex(-1)  # Ничего не выбрано по умолчанию
        self.дробеметная_обработка_специалист.setPlaceholderText("Специалист по дробеметной обработке")
        layout.addWidget(self.дробеметная_обработка_специалист)

        # Заменяем QLineEdit на QComboBox для специалиста по зачистке короны
        self.зачистка_корона_специалист = QComboBox(self)
        self.зачистка_корона_специалист.addItems(zachistka)  # Добавляем участников в комбобокс
        self.зачистка_корона_специалист.setFont(QtGui.QFont("Aptos", 10, QtGui.QFont.Bold))
        self.зачистка_корона_специалист.setStyleSheet("color: black;")
        self.зачистка_корона_специалист.setCurrentIndex(-1)  # Ничего не выбрано по умолчанию
        self.зачистка_корона_специалист.setPlaceholderText("Специалист по зачистке короны")
        layout.addWidget(self.зачистка_корона_специалист)

        # Заменяем QLineEdit на QComboBox для специалиста по зачистке лапы
        self.зачистка_лапа_специалист = QComboBox(self)
        self.зачистка_лапа_специалист.addItems(zachistka)  # Добавляем участников в комбобокс
        self.зачистка_лапа_специалист.setFont(QtGui.QFont("Aptos", 10, QtGui.QFont.Bold))
        self.зачистка_лапа_специалист.setStyleSheet("color: black;")
        self.зачистка_лапа_специалист.setCurrentIndex(-1)  # Ничего не выбрано по умолчанию
        self.зачистка_лапа_специалист.setPlaceholderText("Специалист по зачистке лапы")
        layout.addWidget(self.зачистка_лапа_специалист)

        # Заменяем QLineEdit на QComboBox для специалиста по зачистке питателя
        self.зачистка_питатель_специалист = QComboBox(self)
        self.зачистка_питатель_специалист.addItems(zachistka)  # Добавляем участников в комбобокс
        self.зачистка_питатель_специалист.setFont(QtGui.QFont("Aptos", 10, QtGui.QFont.Bold))
        self.зачистка_питатель_специалист.setStyleSheet("color: black;")
        self.зачистка_питатель_специалист.setCurrentIndex(-1)  # Ничего не выбрано по умолчанию
        self.зачистка_питатель_специалист.setPlaceholderText("Специалист по зачистке питателя")
        layout.addWidget(self.зачистка_питатель_специалист)


        self.примечание = QLineEdit(self)
        self.примечание.setPlaceholderText("Примечание")
        layout.addWidget(self.примечание)

        self.save_button = QPushButton("Сохранить", self)
        self.save_button.setStyleSheet("background-color: #4CAF50; color: black; padding: 10px;")
        self.save_button.setFont(QFont("Arial", 16, QFont.Bold))
        self.save_button.clicked.connect(self.save_data)
        layout.addWidget(self.save_button)

        self.setLayout(layout)

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
        self.номер_кластера.clear()
        self.количество_кластеров.clear()
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
