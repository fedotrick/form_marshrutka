import os
import json
import sys
from datetime import datetime

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.spinner import Spinner
from kivy.uix.popup import Popup
from kivy.core.window import Window
from kivy.properties import ObjectProperty, StringProperty, ListProperty, BooleanProperty
from kivy.uix.dropdown import DropDown
from kivy.graphics import Color, Rectangle, Line, RoundedRectangle
from kivy.metrics import dp
from kivy.utils import get_color_from_hex
from kivy.clock import Clock
from kivy.uix.screenmanager import ScreenManager, Screen, SlideTransition
from kivy.uix.togglebutton import ToggleButton

from openpyxl import Workbook, load_workbook
from functools import partial

# Путь к файлу с данными специалистов
SPECIALISTS_FILE = "specialists.json"

# Функции для загрузки и сохранения списков специалистов
def load_specialists():
    """Загружает списки специалистов из JSON файла"""
    default_specialists = {
        "scleyks": ["Буцик", "Минакова", "Ротарь", "Чернова", "Чупахина"],
        "controlers": ["Елхова", "Шестункина", "Романцева"],
        "bolgar": [
            "Ахмаджонов", "Отаназаров", "Косимов", "Косимов-2", "Туичев",
            "Машрапов", "Эргашев", "Самиев", "Исмаилов"
        ],
        "termob": ["Эгамов", "Аюбов"],
        "drobem": ["Эгамов", "Аюбов"],
        "zachistka": ["Абдуллаев", "Бурхонов", "Матесаев", "Мещерякова",
            "Самиев", "Леонтьева"]
    }
    
    try:
        if os.path.exists(SPECIALISTS_FILE):
            with open(SPECIALISTS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        else:
            # Если файл не существует, создаем его с дефолтными значениями
            save_specialists(default_specialists)
            return default_specialists
    except Exception as e:
        print(f"Ошибка при загрузке списков специалистов: {e}")
        return default_specialists

def save_specialists(specialists_data):
    """Сохраняет списки специалистов в JSON файл"""
    try:
        with open(SPECIALISTS_FILE, 'w', encoding='utf-8') as f:
            json.dump(specialists_data, f, ensure_ascii=False, indent=4)
        return True
    except Exception as e:
        print(f"Ошибка при сохранении списков специалистов: {e}")
        return False

def add_specialist(category, name):
    """Добавляет нового специалиста в указанную категорию"""
    specialists = load_specialists()
    if category in specialists:
        if name not in specialists[category]:
            specialists[category].append(name)
            specialists[category].sort()  # Сортируем список
            save_specialists(specialists)
            return True
    return False

# Функция для загрузки учетных номеров из Excel
def load_account_numbers(file_name):
    if not os.path.exists(file_name):
        return []
        
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
                  дробеметная_обработка_специалист, зачистка_корона_специалист,
                  зачистка_лапа_специалист, зачистка_питатель_специалист, примечание):
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
                    дробеметная_обработка_специалист, зачистка_корона_специалист,
                    зачистка_лапа_специалист, зачистка_питатель_специалист, примечание]
            
            for col, value in enumerate(data, start=1):
                cell = ws.cell(row=next_row, column=col)
                cell.value = value
                
                if col in [1, 4, 10]:  # Колонки с датами
                    cell.number_format = 'DD.MM.YYYY'
                elif col == 5:  # Колонка с временем
                    cell.number_format = 'HH:MM'
            
            wb.save('marshrutka.xlsx')
            wb.close()
            return True
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
                    дробеметная_обработка_специалист, зачистка_корона_специалист,
                    зачистка_лапа_специалист, зачистка_питатель_специалист, примечание]
            
            for col, value in enumerate(data, start=1):
                cell = ws.cell(row=2, column=col)
                cell.value = value
                
                if col in [1, 4, 10]:  # Колонки с датами
                    cell.number_format = 'DD.MM.YYYY'
                elif col == 5:  # Колонка с временем
                    cell.number_format = 'HH:MM'
            
            wb.save('marshrutka.xlsx')
            wb.close()
            return True
    except Exception as e:
        print(f"Ошибка при сохранении в Excel: {str(e)}")
        return False

# Диалог для добавления специалиста
class AddSpecialistPopup(Popup):
    def __init__(self, category, callback, **kwargs):
        super(AddSpecialistPopup, self).__init__(**kwargs)
        self.title = 'Добавление специалиста'
        self.size_hint = (0.8, 0.4)
        self.auto_dismiss = False
        
        self.category = category
        self.callback = callback
        
        layout = BoxLayout(orientation='vertical', padding=20, spacing=15)
        
        # Заголовок
        self.title_label = Label(
            text=f'Введите ФИО специалиста ({self.get_category_name(category)})',
            size_hint_y=None, 
            height=40,
            font_size=18,
            bold=True,
            halign='center'
        )
        layout.add_widget(self.title_label)
        
        # Поле ввода
        self.name_input = TextInput(
            multiline=False,
            size_hint_y=None,
            height=50,
            hint_text='Фамилия специалиста',
            font_size=16,
            padding=[15, 15, 15, 15]
        )
        layout.add_widget(self.name_input)
        
        # Кнопки
        buttons = BoxLayout(size_hint_y=None, height=60, spacing=15)
        
        cancel_button = Button(
            text='Отмена',
            background_color=get_color_from_hex('#DDDBDA'),
            color=get_color_from_hex('#181818')
        )
        cancel_button.bind(on_release=self.dismiss)
        
        add_button = Button(
            text='Добавить',
            background_color=get_color_from_hex('#0176D3'),
            color=get_color_from_hex('#FFFFFF')
        )
        add_button.bind(on_release=self.on_add)
        
        buttons.add_widget(cancel_button)
        buttons.add_widget(add_button)
        layout.add_widget(buttons)
        
        self.content = layout
        
        # Устанавливаем фокус на поле ввода
        Clock.schedule_once(lambda dt: self.name_input.focus, 0.1)
    
    def get_category_name(self, category):
        """Возвращает русское название категории"""
        categories = {
            'scleyks': 'сборка кластера',
            'controlers': 'контроль сборки',
            'bolgar': 'болгарка',
            'termob': 'термообработка',
            'drobem': 'дробеметка',
            'zachistka': 'зачистка'
        }
        return categories.get(category, category)
    
    def on_add(self, instance):
        name = self.name_input.text.strip()
        if not name:
            info_popup = InfoPopup(
                title='Ошибка',
                message='Имя специалиста не может быть пустым',
                size_hint=(0.7, 0.3),
                auto_dismiss=True
            )
            info_popup.open()
            return
        
        success = add_specialist(self.category, name)
        if success:
            self.callback(self.category, name)
            self.dismiss()
        else:
            info_popup = InfoPopup(
                title='Ошибка',
                content=Label(text=f'Не удалось добавить специалиста "{name}". Возможно, он уже существует.'),
                size_hint=(0.7, 0.3)
            )
            info_popup.open()

# Информационный попап
class InfoPopup(Popup):
    def __init__(self, title, message, **kwargs):
        super(InfoPopup, self).__init__(**kwargs)
        self.title = title
        self.size_hint = (0.8, 0.3)
        self.auto_dismiss = True
        
        content = BoxLayout(orientation='vertical', padding=20, spacing=10)
        
        msg_label = Label(
            text=message,
            halign='center',
            font_size=16
        )
        msg_label.bind(size=lambda s, w: setattr(msg_label, 'text_size', (w[0], None)))
        
        button = Button(
            text='OK',
            size_hint=(None, None),
            size=(120, 50),
            pos_hint={'center_x': 0.5},
            background_color=get_color_from_hex('#0176D3'),
            color=get_color_from_hex('#FFFFFF'),
            font_size=16
        )
        button.bind(on_release=self.dismiss)
        
        content.add_widget(msg_label)
        content.add_widget(button)
        self.content = content

# Базовый класс для Контекстного меню
class CustomContextMenu(DropDown):
    def __init__(self, owner=None, **kwargs):
        kwargs.setdefault('auto_width', False)
        kwargs.setdefault('width', 250)
        super(CustomContextMenu, self).__init__(**kwargs)
        self.owner = owner
        
        # Добавляем визуальное оформление
        with self.canvas.before:
            Color(rgba=get_color_from_hex('#FFFFFF'))
            self.rect = Rectangle(pos=self.pos, size=self.size)
            Color(rgba=get_color_from_hex('#DDDBDA'))
            self.border = Line(rectangle=(self.x, self.y, self.width, self.height), width=1)
        
        # Обновляем графику при изменении размера
        self.bind(pos=self.update_graphics, size=self.update_graphics)
    
    def update_graphics(self, instance, value):
        """Обновляет графические элементы при изменении размера"""
        self.rect.pos = self.pos
        self.rect.size = self.size
        self.border.rectangle = (self.x, self.y, self.width, self.height)

class GroupBox(BoxLayout):
    """Специальный класс для GroupBox, чтобы легче находить их при смене темы"""
    is_group_box = BooleanProperty(True)

# Класс для экрана маршрутной карты
class MarshrutkaScreen(BoxLayout):
    # Свойства для доступа к виджетам из KV-файла
    uchet_nomer = ObjectProperty(None)
    naimenovanie = ObjectProperty(None)
    tip_exp = ObjectProperty(None)
    primechanie = ObjectProperty(None)
    
    # Сборка кластера
    sborka_data = ObjectProperty(None)
    sborka_specialist = ObjectProperty(None)
    sborka_kolichestvo = ObjectProperty(None)
    
    # Контроль сборки кластера
    kontrol_data = ObjectProperty(None)
    kontrol_vremya = ObjectProperty(None)
    kontrol_specialist = ObjectProperty(None)
    
    # Обработка
    bolgarka_data = ObjectProperty(None)
    bolgarka_specialist = ObjectProperty(None)
    termob_specialist = ObjectProperty(None)
    drobemet_specialist = ObjectProperty(None)
    
    # Зачистка
    zachistka_korona = ObjectProperty(None)
    zachistka_lapa = ObjectProperty(None)
    zachistka_pitatel = ObjectProperty(None)
    
    # Категории специалистов и соответствующие им виджеты
    specialists_mapping = {
        'scleyks': ['sborka_specialist'],
        'controlers': ['kontrol_specialist'],
        'bolgar': ['bolgarka_specialist'],
        'termob': ['termob_specialist'],
        'drobem': ['drobemet_specialist'],
        'zachistka': ['zachistka_korona', 'zachistka_lapa', 'zachistka_pitatel']
    }
    
    # Порядок вкладок для навигации
    tab_order = ['general', 'assembly', 'processing', 'cleaning']
    
    def __init__(self, **kwargs):
        super(MarshrutkaScreen, self).__init__(**kwargs)
        # Загрузим данные после инициализации всех виджетов
        Clock.schedule_once(self.post_init, 0)
    
    def post_init(self, dt):
        """Инициализация после создания виджетов"""
        # Установка начальной вкладки
        self.current_tab_index = 0
        
        # Настройка перехода между экранами
        self.ids.screen_manager.transition = SlideTransition()
        
        # Загружаем списки специалистов для всех выпадающих списков
        self.load_specialists_data()
        
        # Загружаем учетные номера для выпадающего списка
        self.load_account_numbers()
        
        # Привязываем события к виджетам
        self.setup_event_bindings()
    
    def setup_event_bindings(self):
        """Настройка обработчиков событий для виджетов"""
        # Привязываем обработчик изменения учетного номера
        self.uchet_nomer.bind(text=self.on_account_number_change)
        
        # Привязываем контекстное меню к спиннерам со специалистами
        for category, widget_names in self.specialists_mapping.items():
            for widget_name in widget_names:
                try:
                    widget = getattr(self, widget_name)
                    if widget:
                        # Привязываем обработчик для контекстного меню
                        widget.bind(on_touch_down=partial(self.on_spinner_touch, category))
                        
                        # Добавляем подсказку в заголовок Spinner-а 
                        # Это поможет пользователю понять, что можно добавить специалиста
                        spinner_text = widget.text or "Выберите или добавьте специалиста (правый клик)"
                        widget.text = spinner_text
                except Exception as e:
                    print(f"Ошибка при привязке обработчика к {widget_name}: {e}")
    
    def switch_tab(self, tab_name):
        """Переключает на указанную вкладку"""
        if tab_name in self.tab_order:
            self.current_tab_index = self.tab_order.index(tab_name)
            # Устанавливаем направление перехода для ScreenManager
            self.ids.screen_manager.transition.direction = 'left'
            self.ids.screen_manager.current = tab_name
    
    def next_tab(self):
        """Переход к следующей вкладке"""
        if self.current_tab_index < len(self.tab_order) - 1:
            self.current_tab_index += 1
            self.ids.screen_manager.transition.direction = 'left'
            self.ids.screen_manager.current = self.tab_order[self.current_tab_index]
    
    def previous_tab(self):
        """Переход к предыдущей вкладке"""
        if self.current_tab_index > 0:
            self.current_tab_index -= 1
            self.ids.screen_manager.transition.direction = 'right'
            self.ids.screen_manager.current = self.tab_order[self.current_tab_index]
    
    def get_current_date(self):
        """Возвращает текущую дату в формате ДД.ММ.ГГГГ"""
        return datetime.now().strftime('%d.%m.%Y')
    
    def get_current_time(self):
        """Возвращает текущее время в формате ЧЧ:ММ"""
        return datetime.now().strftime('%H:%M')
    
    def on_spinner_touch(self, category, instance, touch):
        """Обработка нажатия на спиннер для вызова контекстного меню"""
        if not instance.collide_point(*touch.pos):
            return False
            
        # Проверяем, что это правый клик (кнопка 3)
        if touch.button == 'right':
            # Предотвращаем дальнейшую обработку события
            touch.ud['handled'] = True
            
            # Создаем и открываем контекстное меню
            self.show_context_menu(instance, category)
            return True
        return False
    
    def show_context_menu(self, spinner, category):
        """Показывает контекстное меню для добавления нового специалиста"""
        menu = CustomContextMenu(spinner)
        
        # Добавляем пункт меню
        item = Button(
            text='Добавить специалиста', 
            size_hint_y=None, 
            height=44,
            background_color=get_color_from_hex('#0176D3'),
            color=get_color_from_hex('#FFFFFF')
        )
        item.bind(on_release=lambda btn: self.show_add_specialist_popup(category, menu))
        menu.add_widget(item)
        
        # Открываем меню
        menu.open(spinner)
    
    def show_add_specialist_popup(self, category, menu=None):
        """Показывает диалог для добавления нового специалиста"""
        # Закрываем контекстное меню, если оно открыто
        if menu:
            menu.dismiss()
            
        popup = AddSpecialistPopup(
            category=category,
            callback=self.on_specialist_added
        )
        popup.open()
    
    def on_specialist_added(self, category, name):
        """Обработчик события добавления нового специалиста"""
        # Обновляем соответствующие выпадающие списки
        self.update_specialists_spinners(category)
        
        # Устанавливаем добавленного специалиста как выбранного в соответствующих спиннерах
        if category in self.specialists_mapping:
            for widget_name in self.specialists_mapping[category]:
                widget = getattr(self, widget_name)
                # Используем новое значение, если спиннер был пустым или содержал подсказку
                current_text = widget.text or ""
                if not current_text or "Выберите или добавьте" in current_text:
                    widget.text = name
        
        # Показываем сообщение об успешном добавлении
        info_popup = InfoPopup(
            title='Специалист добавлен',
            message=f'Специалист "{name}" успешно добавлен и доступен во всех соответствующих полях'
        )
        info_popup.open()
    
    def update_specialists_spinners(self, category):
        """Обновляет выпадающие списки для определенной категории специалистов"""
        specialists = load_specialists()
        if category in specialists and category in self.specialists_mapping:
            for widget_name in self.specialists_mapping[category]:
                widget = getattr(self, widget_name)
                
                # Запоминаем текущее значение
                current_value = widget.text
                
                # Обновляем список значений
                widget.values = specialists[category]
                
                # Восстанавливаем выбранное значение, если оно есть в обновленном списке
                if current_value in specialists[category]:
                    widget.text = current_value
    
    def load_specialists_data(self):
        """Загружает данные о специалистах во все выпадающие списки"""
        specialists = load_specialists()
        
        # Заполняем все выпадающие списки
        for category, widget_names in self.specialists_mapping.items():
            if category in specialists:
                for widget_name in widget_names:
                    widget = getattr(self, widget_name)
                    widget.values = specialists[category]
    
    def load_account_numbers(self):
        """Загружает учетные номера в выпадающий список"""
        try:
            account_numbers = load_account_numbers('plavka.xlsx')
            self.uchet_nomer.values = account_numbers
        except Exception as e:
            print(f"Ошибка при загрузке учетных номеров: {e}")
    
    def on_account_number_change(self, instance, value):
        """Обработчик изменения учетного номера"""
        if not value:
            return
            
        try:
            if os.path.exists('plavka.xlsx'):
                wb = load_workbook('plavka.xlsx')
                sheet = wb.active
                
                # Ищем строку с выбранным учетным номером
                for row in sheet.iter_rows(min_row=2, values_only=True):
                    if row[1] == value:  # Учетный_номер во втором столбце
                        # Наименование отливки в 11-м столбце (индекс 10)
                        if row[10]:  # Наименование_отливки
                            self.naimenovanie.text = str(row[10])
                        
                        # Тип эксперимента в 12-м столбце (индекс 11)
                        if row[11]:  # Тип_эксперемента
                            self.tip_exp.text = str(row[11])
                        
                        # Дата плавки в 3-м столбце (индекс 2)
                        плавка_дата = row[2]  # Плавка_дата
                        
                        if isinstance(плавка_дата, datetime):
                            # Если дата в формате datetime
                            self.bolgarka_data.text = плавка_дата.strftime("%d.%m.%Y")
                        elif isinstance(плавка_дата, str):
                            # Если дата в строковом формате DD.MM.YYYY
                            self.bolgarka_data.text = плавка_дата
                        break
                
                wb.close()
        except Exception as e:
            print(f"Ошибка при обновлении данных: {e}")
    
    def on_spinner_select(self, spinner, text):
        """Обработчик выбора элемента в выпадающем списке"""
        pass
    
    def validate_time(self, time_str):
        """Проверка корректности ввода времени в формате ЧЧ:ММ"""
        try:
            hours, minutes = map(int, time_str.split(':'))
            if 0 <= hours < 24 and 0 <= minutes < 60:
                return True
        except ValueError:
            return False
        return False
    
    def validate_date(self, date_str):
        """Проверка корректности ввода даты в формате ДД.ММ.ГГГГ"""
        try:
            day, month, year = map(int, date_str.split('.'))
            # Базовая проверка без учета високосных лет и разного количества дней в месяцах
            if 1 <= day <= 31 and 1 <= month <= 12 and 1900 <= year <= 2100:
                return True
        except ValueError:
            return False
        return False
    
    def save_data(self):
        """Сохраняет данные формы в Excel-файл"""
        # Проверяем обязательные поля
        required_fields = [
            (self.uchet_nomer, "Учетный номер"),
            (self.naimenovanie, "Наименование отливки"),
            (self.tip_exp, "Тип эксперимента"),
            (self.sborka_specialist, "Специалист сборки"),
            (self.kontrol_specialist, "Специалист контроля")
        ]
        
        empty_fields = []
        for field, name in required_fields:
            if not field.text:
                empty_fields.append(name)
        
        if empty_fields:
            # Показываем информационное сообщение
            popup = InfoPopup(
                title='Не все поля заполнены',
                message="Пожалуйста, заполните следующие обязательные поля:\n• " + "\n• ".join(empty_fields)
            )
            popup.open()
            
            # Переключаемся на вкладку с первым незаполненным полем
            if "Учетный номер" in empty_fields or "Наименование отливки" in empty_fields or "Тип эксперимента" in empty_fields:
                self.switch_tab('general')
            elif "Специалист сборки" in empty_fields or "Специалист контроля" in empty_fields:
                self.switch_tab('assembly')
            
            return
        
        # Проверяем корректность формата даты
        date_fields = [
            (self.sborka_data, "Дата сборки"),
            (self.kontrol_data, "Дата выставления"),
            (self.bolgarka_data, "Дата болгарки")
        ]
        
        invalid_dates = []
        for field, name in date_fields:
            if field.text and not self.validate_date(field.text):
                invalid_dates.append(name)
        
        if invalid_dates:
            # Показываем информационное сообщение
            popup = InfoPopup(
                title='Неверный формат даты',
                message="Неверный формат даты в полях:\n• " + "\n• ".join(invalid_dates) + "\nИспользуйте формат ДД.ММ.ГГГГ"
            )
            popup.open()
            
            # Переключаемся на вкладку с первым некорректным полем
            if "Дата сборки" in invalid_dates or "Дата выставления" in invalid_dates:
                self.switch_tab('assembly')
            elif "Дата болгарки" in invalid_dates:
                self.switch_tab('processing')
            
            return
        
        # Проверяем корректность формата времени
        if not self.validate_time(self.kontrol_vremya.text):
            popup = InfoPopup(
                title='Неверный формат времени',
                message="Некорректный ввод времени. Используйте формат ЧЧ:ММ."
            )
            popup.open()
            self.switch_tab('assembly')
            return
        
        # Собираем данные из полей формы
        сборка_кластера_дата = self.sborka_data.text
        сборка_кластера_специалист = self.sborka_specialist.text
        сборка_кластера_количество = self.sborka_kolichestvo.text
        контроль_сборки_кластера_дата_выставления = self.kontrol_data.text
        контроль_сборки_кластера_время_выставления = self.kontrol_vremya.text
        контроль_сборки_кластера_специалист = self.kontrol_specialist.text
        учетный_номер = self.uchet_nomer.text
        наименование_отливки = self.naimenovanie.text
        тип_эксперемента = self.tip_exp.text
        болгарка_дата = self.bolgarka_data.text
        болгарка_специалист = self.bolgarka_specialist.text
        термообработка_специалист = self.termob_specialist.text
        дробеметная_обработка_специалист = self.drobemet_specialist.text
        зачистка_корона_специалист = self.zachistka_korona.text
        зачистка_лапа_специалист = self.zachistka_lapa.text
        зачистка_питатель_специалист = self.zachistka_pitatel.text
        примечание = self.primechanie.text
        
        # Сохраняем данные в Excel
        success = save_to_excel(
            сборка_кластера_дата, сборка_кластера_специалист, сборка_кластера_количество,
            контроль_сборки_кластера_дата_выставления, контроль_сборки_кластера_время_выставления,
            контроль_сборки_кластера_специалист, учетный_номер, наименование_отливки, тип_эксперемента,
            болгарка_дата, болгарка_специалист, термообработка_специалист,
            дробеметная_обработка_специалист, зачистка_корона_специалист,
            зачистка_лапа_специалист, зачистка_питатель_специалист, примечание
        )
        
        if success:
            # Показываем сообщение об успешном сохранении
            popup = InfoPopup(
                title='Успех',
                message='Данные успешно сохранены в Excel!'
            )
            popup.open()
            
            # Обновляем список учетных номеров
            self.load_account_numbers()
            
            # Очищаем поля ввода
            self.clear_fields()
            
            # Переключаемся на первую вкладку для нового ввода
            self.switch_tab('general')
        else:
            # Показываем сообщение об ошибке
            popup = InfoPopup(
                title='Ошибка',
                message='Ошибка при сохранении данных в Excel!'
            )
            popup.open()
    
    def clear_fields(self):
        """Очищает все поля формы"""
        # Текстовые поля
        self.naimenovanie.text = ''
        self.tip_exp.text = ''
        self.primechanie.text = ''
        
        # Поля ввода даты
        self.sborka_data.text = self.get_current_date()
        self.kontrol_data.text = self.get_current_date()
        self.bolgarka_data.text = self.get_current_date()
        
        # Поле ввода времени
        self.kontrol_vremya.text = self.get_current_time()
        
        # Количество
        self.sborka_kolichestvo.text = ''
        
        # Выпадающие списки специалистов
        self.sborka_specialist.text = ''
        self.kontrol_specialist.text = ''
        self.bolgarka_specialist.text = ''
        self.termob_specialist.text = ''
        self.drobemet_specialist.text = ''
        self.zachistka_korona.text = ''
        self.zachistka_lapa.text = ''
        self.zachistka_pitatel.text = ''
        
        # Учетный номер
        self.uchet_nomer.text = ''

class MarshrutkaApp(App):
    def build(self):
        self.title = 'Электронная маршрутная карта'
        # Устанавливаем светлую тему по умолчанию
        self.theme_mode = 'light'
        self.theme_button = None  # Для хранения ссылки на кнопку переключения темы
        
        # Привязываем обработчик изменения размера окна
        Window.bind(on_resize=self.on_window_resize)
        
        # Получаем ссылку на корневой виджет
        root = MarshrutkaScreen()
        
        # Отложенное обновление темы после инициализации
        Clock.schedule_once(self.post_init, 0.5)
        
        return root
    
    def post_init(self, dt):
        """Инициализация после построения всех виджетов"""
        try:
            # Получение ссылки на кнопку переключения темы
            self.theme_button = self.root.ids.theme_button
        except Exception as e:
            print(f"Ошибка в post_init: {e}")
    
    def on_window_resize(self, instance, width, height):
        """Обновляет тему при изменении размера окна"""
        # Используем Clock для отложенного обновления после изменения размера
        Clock.schedule_once(lambda dt: self.update_theme_colors(), 0.1)
    
    def toggle_theme(self):
        """Переключает между светлой и темной темой"""
        try:
            self.theme_mode = 'dark' if self.theme_mode == 'light' else 'light'
            
            # Обновляем текст кнопки
            if self.theme_button:
                self.theme_button.text = 'Светлая тема' if self.theme_mode == 'dark' else 'Темная тема'
            
            # Обновляем цвета темы    
            self.update_theme_colors()
        except Exception as e:
            print(f"Ошибка при переключении темы: {e}")
    
    def update_theme_colors(self):
        """Обновляет цвета интерфейса в соответствии с текущей темой"""
        try:
            # Цвета для разных тем
            colors = {
                'light': {
                    'background': '#F3F3F3',
                    'card_background': '#FFFFFF',
                    'text': '#181818',
                    'input_bg': '#FFFFFF',
                    'input_text': '#181818', 
                    'accent': '#0176D3',
                    'border': '#DDDBDA'
                },
                'dark': {
                    'background': '#1A1C1E',
                    'card_background': '#2D2D2D',
                    'text': '#FFFFFF',
                    'input_bg': '#353535',
                    'input_text': '#FFFFFF',
                    'accent': '#0176D3', 
                    'border': '#555555'
                }
            }
            
            theme = colors[self.theme_mode]
            
            # Обновляем глобальный фон корневого виджета
            self.root.canvas.before.clear()
            with self.root.canvas.before:
                Color(rgba=get_color_from_hex(theme['background']))
                Rectangle(pos=self.root.pos, size=self.root.size)
            
            # Функция для обновления GroupBox
            def update_group_box(box):
                try:
                    box.canvas.before.clear()
                    with box.canvas.before:
                        Color(rgba=get_color_from_hex(theme['card_background']))
                        RoundedRectangle(pos=box.pos, size=box.size, radius=[8])
                        Color(rgba=get_color_from_hex(theme['border']))
                        Line(rounded_rectangle=[box.x, box.y, box.width, box.height, 8], width=1)
                except Exception as e:
                    print(f"Ошибка при обновлении GroupBox: {e}")
            
            # Сначала обновляем текущий видимый экран
            current_screen = self.root.ids.screen_manager.current_screen
            self.update_screen_widgets(current_screen, theme, update_group_box)
            
            # Затем обновляем все остальные экраны
            for screen_name in self.root.tab_order:
                if screen_name != self.root.ids.screen_manager.current:
                    screen = self.root.ids.screen_manager.get_screen(screen_name)
                    self.update_screen_widgets(screen, theme, update_group_box)
            
            # Обновляем виджеты корневого экрана (вне ScreenManager)
            self.update_root_widgets(theme)
            
            # Принудительно обновляем виджеты
            self.root.canvas.ask_update()
        except Exception as e:
            print(f"Ошибка при обновлении темы: {e}")
    
    def update_screen_widgets(self, screen, theme, update_group_box):
        """Обновляет виджеты на указанном экране"""
        try:
            if not screen:
                return
                
            for child in screen.walk():
                try:
                    # Обновляем GroupBox
                    if child.__class__.__name__ == 'GroupBox' or 'groupbox' in str(child.__class__).lower():
                        update_group_box(child)
                    
                    # Обновляем CustomLabel
                    elif isinstance(child, Label) and not isinstance(child, Button):
                        # Не меняем цвет акцентных заголовков (GroupTitle)
                        if hasattr(child, 'bold') and child.bold and 'groupt' in str(child.__class__).lower():
                            continue
                        child.color = get_color_from_hex(theme['text'])
                    
                    # Обновляем CustomTextInput
                    elif isinstance(child, TextInput):
                        child.background_color = get_color_from_hex(theme['input_bg'])
                        child.foreground_color = get_color_from_hex(theme['input_text'])
                        child.cursor_color = get_color_from_hex(theme['accent'])
                    
                    # Обновляем CustomSpinner
                    elif isinstance(child, Spinner):
                        child.background_color = get_color_from_hex(theme['input_bg'])
                        child.color = get_color_from_hex(theme['input_text'])
                    
                    # Обновляем TabButton
                    elif isinstance(child, ToggleButton) and hasattr(child, 'group') and child.group == 'tabs':
                        child.background_color = (
                            get_color_from_hex(theme['card_background']) 
                            if child.state == 'normal' 
                            else get_color_from_hex(theme['accent'])
                        )
                        child.color = (
                            get_color_from_hex(theme['accent']) 
                            if child.state == 'normal' 
                            else get_color_from_hex('#FFFFFF')
                        )
                except Exception as e:
                    print(f"Ошибка при обновлении виджета {child}: {e}")
        except Exception as e:
            print(f"Ошибка при обновлении виджетов экрана {screen}: {e}")
    
    def update_root_widgets(self, theme):
        """Обновляет виджеты корневого экрана"""
        try:
            # Обновляем заголовок и кнопки
            for child in self.root.children:
                if isinstance(child, BoxLayout):
                    for widget in child.walk():
                        try:
                            # Обновляем кнопки навигации
                            if (isinstance(widget, Button) and 
                                not isinstance(widget, ToggleButton) and 
                                widget.text in ['← Назад', 'Вперед →']):
                                widget.background_color = get_color_from_hex(theme['card_background'])
                                widget.color = get_color_from_hex(theme['text'])
                            
                            # Обновляем кнопку сохранения (зеленая кнопка)
                            elif (isinstance(widget, Button) and 
                                  not isinstance(widget, ToggleButton) and 
                                  widget.text == 'Сохранить'):
                                pass  # Оставляем зеленую кнопку без изменений
                            
                            # Обновляем кнопку темы
                            elif widget is self.theme_button:
                                widget.background_color = get_color_from_hex(theme['card_background'])
                                widget.color = get_color_from_hex(theme['text'])
                        except Exception as e:
                            print(f"Ошибка при обновлении корневого виджета {widget}: {e}")
        except Exception as e:
            print(f"Ошибка при обновлении корневых виджетов: {e}")

if __name__ == '__main__':
    MarshrutkaApp().run()
