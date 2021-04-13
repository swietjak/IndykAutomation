from pynput import mouse #import Button, Controller
from pynput import keyboard
import time
import openpyxl
import pyperclip
from controller_lib import Base_Clicker

class Clicker(Base_Clicker):
    def __init__(self , letter):
        super().__init__()
        self.adress_coords = (207, 197)
        self.sort_group_coords = (391, 306)
        self.sex_coords = (420, 283)
        self.prep_coords = (576, 598)
    
    def read_sheet(self, wb_filename):
        wb = openpyxl.load_workbook(wb_filename)
        self.sheet = wb['swieze']

    def pick_sex(self, id, i):
        field = self.sheet['L' + str(i)].value
        if field == 'indor':
            return '1'
        elif field == 'indyczka':
            return '2'
        elif field == 'indyk':
            return '3'
    
    def insert_sort_group(self, id, row):
        time.sleep(self.interval)
        self.move_click(self.sort_group_coords)
        time.sleep(self.interval)
        self.move_click(self.prep_coords)
        time.sleep(self.interval)
        self.move_click(self.sex_coords)
        time.sleep(self.interval)
        self.keyboard_controller.type(self.pick_sex(id, row))
        time.sleep(self.interval)
        self.click_kb(keyboard.Key.esc)

    def do_action(self, start, stop):
        for i in range(start, stop):
            if self.sheet['L' + str(i)].value == 'indor':
                continue
            id = self.sheet['A' + str(i)].value
            # print(self.pick_durability(id))
            self.move_click(self.adress_coords)
            time.sleep(self.interval)
            self.keyboard_controller.type(str(id) + "020")
            self.click_kb(keyboard.Key.enter)
            time.sleep(self.interval)
            self.insert_sort_group(id, i)
            time.sleep(self.interval)
            self.click_kb(keyboard.Key.f2)
            time.sleep(2)

cl = Clicker('D')
cl.read_sheet('asort.xlsx')
time.sleep(2)
cl.do_action(54, 66)
