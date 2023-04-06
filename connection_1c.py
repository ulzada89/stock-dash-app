from pywinauto.keyboard import send_keys
from pywinauto.mouse import click, double_click
from pywinauto import Application
import time
import datetime


DATE = '01.01.2023'
TODAY_DATE = datetime.date.today()
PREV_DATE = datetime.datetime.strftime(TODAY_DATE - datetime.timedelta(days=1), '%d.%m.%Y')
FIRST_DAY_OF_MONTH = f"01.{datetime.datetime.strftime(datetime.date.today(), '%m.%Y')}"
MENU_TABS = ['Операции', 'Отгрузка', 'Отгрузка']
MENU_ITEMS = ['Отчеты по операциям', 'Отчеты по отгрузке', 'Отчеты по отгрузке']
REPORT_ITEMS = ['Stock Level - for Marketing', 'Polymer Shipment - Local', 'Polymer Shipment - Export']
COORDS = [[(415, 102), (567, 102), (679, 104), (373, 173), (887, 171)],
          [(521, 102), (648, 102), (758, 104), (323, 198), (888, 201)],
          [(521, 102), (648, 102), (758, 104), (323, 198), (888, 201)]]


class Start1C:
    def __init__(self):
        Application(backend="uia").start(r'C:/Program Files/1cv8/common/1cestart.exe', timeout=60)
        self.app = Application(backend="uia").connect(title_re='.*Предпри.*', timeout=60)
        self.app.Dialog.ListItem3.double_click_input()
        self.dlg = Application(backend='uia').connect(title_re='.*Доступ.*', timeout=60)
        self.dlg.Dialog.Edit.type_keys('point0990')
        self.dlg.Dialog.Button2.click_input()

    def connect_to_main_window(self):
        for i in range(3):
            self.one_c_main_window = Application(backend='uia').connect(title_re='.*Логистика..*', timeout=60)
            self.one_c_main_window = self.one_c_main_window.window(title_re='.*Логистика..*')
            self.one_c_main_window.set_focus()
            self.one_c_main_window.child_window(title=MENU_TABS[i], control_type="TabItem").click_input()
            time.sleep(1)
            self.one_c_main_window.child_window(title=MENU_ITEMS[i], control_type="MenuItem").click_input()
            time.sleep(1)
            self.one_c_main_window.child_window(title=REPORT_ITEMS[i], control_type="Hyperlink").double_click_input()
            send_keys('{ENTER}')
            time.sleep(1)
            self.one_c_main_window = Application(backend='uia').connect(title_re='.*Логистика..*', timeout=60)
            self.one_c_main_window.window(title_re='.*Логистика..*').set_focus()
            double_click(coords=COORDS[i][0])
            time.sleep(1)
            send_keys('Произвольный')
            double_click(coords=COORDS[i][1])
            if i == 1:
                send_keys(DATE)
            else:
                send_keys(FIRST_DAY_OF_MONTH)
            double_click(coords=COORDS[i][2])
            send_keys(PREV_DATE)
            click(coords=COORDS[i][3])
            time.sleep(10)
            click(coords=COORDS[i][4])
            time.sleep(3)
            self.save_dlg = Application().connect(title_re='.*Сохранить.*', timeout=60)
            self.save_dlg.Dialog.AddressBandRoot.click_input()
            self.save_dlg.Dialog.AddressBandRoot.type_keys(r'C:\Users\u.kutlimuratova\Desktop\DAILY STOCK LEVEL', with_spaces=True)
            send_keys('{ENTER}')
            self.save_dlg.Dialog.ComboBox2.select("Лист Excel (*.xls)")
            self.save_dlg.Dialog.Edit.type_keys(REPORT_ITEMS[i], with_spaces=True)
            self.save_dlg.Dialog.Button.click_input()
            try:
                self.save_as_dlg = Application().connect(title_re='.*сохранение.*', timeout=60)
                self.save_as_dlg.Dialog.set_focus()
                self.save_as_dlg.Dialog.Button.click_input()
            except:
                pass



