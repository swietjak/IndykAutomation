from pynput import mouse #import Button, Controller
from pynput import keyboard
import time
import openpyxl
import pyperclip
import math
from controller_lib import click_kb, move_click

class Clicker():
    def __init__(self , letter):
        self.adress_coords = (545, 164)
        self.rules_coords = (813, 233)

        self.mouse_controller = mouse.Controller()
        self.keyboard_controller = keyboard.Controller()
        self.interval = 0.2

    def read_sheet(self, wb_filename):
        wb = openpyxl.load_workbook(wb_filename)
        self.sheet = wb['page 1']

    def pick_lang(self, lang, id):
        if lang == None:
            return None
        if lang == "CZ":
            return "99850"
        elif lang == "SK":
            return "99854"
        elif lang == "FR":
            return "99040"
        elif lang == "D" or lang == "A":
            return "99030"
        elif lang == "ES":
            return "99060"
        else:
            with open("txt_files/langs.txt", 'a') as f:
                f.write(id + " " + lang)
            return "99020"
        
    def validate_id(self, id):
        if id == None:
            return False            
        try:
            id = int(id)
        except ValueError:
            return False
        
        return True

    def validate_lang(self, lang, i):
        if lang == None:
            lang = self.sheet['G' + str(i+1)].value
            if lang == None:
                return False, None
        if len(lang) > 5:
            lang = self.sheet['G' + str(i+1)].value
            if lang == None:
                return False, None
        if lang == "PL":
                return False, None
        
        return True, lang

    def do_action(self, start, stop):
        self.langs = []
        for i in range(start, stop):
            id = self.sheet['C' + str(i)].value
            lang = self.sheet['H' + str(i+1)].value
            lang_check, lang = self.validate_lang(lang, i)
            if self.validate_id(id) == False or lang_check == False :
                continue
            
            l = self.pick_lang(lang, id)
            # print(lang)
            print(i)
            move_click(self.mouse_controller, self.adress_coords)
            self.keyboard_controller.type(str(id))
            time.sleep(self.interval)
            click_kb(self.keyboard_controller, keyboard.Key.enter)
            time.sleep(self.interval)
            move_click(self.mouse_controller, self.rules_coords)
            time.sleep(3)
            click_kb(self.keyboard_controller, keyboard.Key.tab)
            self.keyboard_controller.type(l)
            time.sleep(self.interval)
            click_kb(self.keyboard_controller, keyboard.Key.f2)
            print(f"Zrobiono klienta {id} jezyk {lang}")
            time.sleep(2)
cl = Clicker('E')
cl.read_sheet('./sheets/full_addr_list.xlsx')    
cl.do_action(14620, 15818)