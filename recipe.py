import controller_lib as clib
import time
from pynput import mouse #import Button, Controller
from pynput import keyboard
import openpyxl
import pyperclip


class Clicker(clib.Base_Clicker):
    def __init__(self):
        super().__init__()
        self.addr_cords = self.convert_coords((278, 256))
        self.data_ii =  self.convert_coords((213, 381))

    def change_storage(self, id):
        self.move_click(self.addr_cords)
        self.keyboard_controller.type(str(id))
        time.sleep(self.interval)
        self.click_kb(keyboard.Key.enter)
        time.sleep(self.interval)
        self.move_click(self.data_ii)
        time.sleep(self.interval)
        self.click_kb(keyboard.Key.down, 2)
        self.keyboard_controller.type('261')
        time.sleep(self.interval)
        self.click_kb(keyboard.Key.f2)
        time.sleep(self.interval)

    def do_action(self, start, stop):
        for i in range(start, stop):
            id = self.sheet['A' + str(i)].value
            if (int(id) - 90) % 100 == 0:
                self.change_storage(id)
                time.sleep(1)

cl = Clicker()
cl.read_sheet("sheets/asort.xlsx", "swieze")
time.sleep(2)
cl.do_action(10, 695)
cl.change_storage("11090")