import time
from pynput import mouse #import Button, Controller
from pynput import keyboard
import openpyxl
import pyperclip
import numpy as np

class Base_Clicker():
    def __init__(self):
        self.mouse_controller = mouse.Controller()
        self.keyboard_controller = keyboard.Controller()
        self.interval = 0.3
    
    def click_kb(self, key, times=1):
        for _ in range(times):
            self.keyboard_controller.press(key)
            self.keyboard_controller.release(key)

    def move_click(self, pos, interval=0.1):
        self.mouse_controller.position = pos
        time.sleep(interval)
        self.mouse_controller.click(mouse.Button.left)

    def drag(self, start_pos, end_pos):
        self.mouse_controller.position = start_pos
        time.sleep(0.1)
        self.mouse_controller.press(mouse.Button.left)
        self.mouse_controller.position = end_pos
        self.mouse_controller.release(mouse.Button.left)
    
    def read_sheet(self, wb_filename, sheet_name):
        wb = openpyxl.load_workbook(wb_filename)
        self.sheet = wb[sheet_name]
    
    def do_action(self):
        raise NotImplementedError()

    def convert_coords(self, coords, ratio=1.25):
        return tuple(np.array(coords) / ratio)