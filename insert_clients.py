from pynput import mouse #import Button, Controller
from pynput import keyboard
import time
import openpyxl
import pyperclip
from controller_lib import Base_Clicker

class Clicker(Base_Clicker):
    def __init__(self):
        super().__init__()
        self.adress_coords = self.convert_coords((238, 143))
        self.lang_coords = self.convert_coords((181, 212))

    def insert_client(self, index):
        name = self.sheet['B' + str(index)].value
        self.move_click(self.adress_coords)
        self.keyboard_controller.type(str(index) + '\n')
        self.keyboard_controller.type(name + '\t')
        self.keyboard_controller.type(name + '\t')
        self.keyboard_controller.type('11')
        self.keyboard_controller.type(name + '\t')
        self.click_kb(keyboard.Key.tab, 15)
        self.keyboard_controller.type('1')
        self.move_click(self.lang_coords)
        self.click_kb(keyboard.Key.down)
        self.keyboard_controller.type('84')

    def do_action(self, start, stop, index):
        for i in range(start, stop):
            if self.sheet['C' + str(i)].value:
                self.insert_client(i)
                    

cl = Clicker()
cl.read_sheet("klienci.xlsx", "Arkusz1")
time.sleep(2)
cl.do_action(17, 18, 7322)
