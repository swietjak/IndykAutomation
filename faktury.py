from pynput import mouse #import Button, Controller
from pynput import keyboard
import time
import openpyxl
import pyperclip
from controller_lib import Base_Clicker

class Clicker(Base_Clicker):
    def __init__(self):
        self.adress_coords = (175, 202)
        self.docs_coords = (252, 265)

    def insert_receit(self, id, lang_bool):
        wz_value = '2'
        receit_value = '1' if lang_bool else '4'
        self.move_click(self.adress_coords)
        time.sleep(self.interval)
        self.keyboard_controller.type(str(id))
        self.click_kb(keyboard.Key.enter)
        time.sleep(self.interval)
        #nawigacja do zamówienia
        self.click_kb(keyboard.Key.down, 2)
        self.click_kb(keyboard.Key.enter)
        time.sleep(self.interval)
        self.click_kb(keyboard.Key.down)
        self.click_kb(keyboard.Key.enter)

        time.sleep(3*self.interval)
        self.move_click(self.docs_coords)
        time.sleep(3)
        self.click_kb(keyboard.Key.down, 5)
        self.keyboard_controller.type(wz_value + '\t')
        self.keyboard_controller.type(receit_value + '\t')
        time.sleep(self.interval)
        self.click_kb(keyboard.Key.f2)
        time.sleep(self.interval)
        self.click_kb(keyboard.Key.esc)

    def do_action(self, start, stop):
        for i in range(start, stop):
            id = self.sheet['A' + str(i)].value
            if id is not None and int(id) > 1000:
                lang_bool = True if self.sheet['B' + str(i)].value == "pl" else False
                self.insert_receit(id, lang_bool)
                with open("log.txt", 'a') as f:
                    message = f'Zrobiono klienta nr: {id}, rzad: {i}, progres: {str(i * 100/self.sheet.max_row)}%\n'
                    f.write(message)
                    print(message)
                time.sleep(self.interval)

cl = Clicker()
cl.read_sheet("dod_klienci.xlsx", 'Arkusz1')
time.sleep(1.5)
cl.do_action(1, 126)
