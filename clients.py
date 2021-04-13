from pynput import mouse #import Button, Controller
from pynput import keyboard
import time
import openpyxl
import pyperclip

def click_kb(controller, key, times=1):
    for i in range(times):
        controller.press(key)
        controller.release(key)

def move_click(controler, pos, interval=0.1):
    controler.position = pos
    time.sleep(interval)
    controler.click(mouse.Button.left)

class Clicker():
    def __init__(self):
        self.adress_coords = (248, 181)
        self.price_coords = (109, 244)
        self.const_price_coords = (313, 544)
        self.pay_rules_set_coords = (335, 455)
        self.pay_rules_coords = (173, 246)
        self.next_coords = (601, 343)
        self.docs_coords = (253, 245)
        self.etyciet_coords = (309, 240)

        self.mouse_controller = mouse.Controller()
        self.keyboard_controller = keyboard.Controller()
        self.interval = 0.2

    def read_sheet(self, wb_filename):
        wb = openpyxl.load_workbook(wb_filename)
        self.sheet = wb['Arkusz1']

    def return_lang_num(self, lang):
        num = 0
        if(lang == 'ang'):
            return '113'
        elif lang == 'cz':
            return '112'
        elif lang == 'hsp':
            return '115'
        elif lang == 'fr':
            return '116'
        elif lang == 'niem':
            return '114'
        elif lang == 'gre':
            return '118'
        elif lang == 'rus':
            return '119'
        elif lang == 'chr':
            return '120'
        elif lang == 'por':
            return '121'
        elif lang == 'rum':
            return '122'
        elif lang == 'ukr':
            return '123'
        elif lang == 'slo':
            return '124'
        elif lang == 'wlo':
            return '125'
        elif lang == 'nid':
            return '126'
        else:
            return '111'

    def do_action(self, start, stop):
        for i in range(start, stop):
            id = self.sheet['A' + str(i)].value
            lang = self.sheet['B' + str(i)].value
            print(self.return_lang_num(lang))
            if id is None:
                continue
            else:
                print(str(id) + ' ' + str(i))
                self.insert_order(int(id), lang)
                time.sleep(4*self.interval)
                with open('log.txt', 'a') as f:
                    f.write(f'Zrobiono klienta nr: {id}, rzad: {i}\n')
                    print(f'Zrobiono klienta nr: {id}, rzad: {i}, progres: {str(i * 100/self.sheet.max_row)}%\n')

    def insert_order(self, number, lang):
        #wpisywanie numeru
        self.mouse_controller.position = self.adress_coords
        self.mouse_controller.click(mouse.Button.left)
        #(795, 324)
        time.sleep(self.interval)
        self.keyboard_controller.type(str(number))
        click_kb(self.keyboard_controller, keyboard.Key.enter)
        # move_click(self.mouse_controller, self.next_coords)
        time.sleep(self.interval)
        #nawigacja do zamówienia
        click_kb(self.keyboard_controller, keyboard.Key.down)
        click_kb(self.keyboard_controller, keyboard.Key.down)
        # click_kb(self.keyboard_controller, keyboard.Key.down)
        click_kb(self.keyboard_controller, keyboard.Key.enter)
        time.sleep(1.5)
        click_kb(self.keyboard_controller, keyboard.Key.down)
        click_kb(self.keyboard_controller, keyboard.Key.enter)

        time.sleep(self.interval)
        self.insert_ord()
        time.sleep(self.interval)
        self.insert_price()
        time.sleep(self.interval)
        self.insert_payrules()
        time.sleep(self.interval)
        self.insert_docs()
        time.sleep(self.interval)
        self.insert_etyciet(lang)

        time.sleep(self.interval)
        click_kb(self.keyboard_controller, keyboard.Key.f2)
        time.sleep(self.interval)
        click_kb(self.keyboard_controller, keyboard.Key.esc)
        time.sleep(self.interval)
    
    def insert_ord(self):
        self.keyboard_controller.type('10')#rodz wys
        click_kb(self.keyboard_controller, keyboard.Key.down, 2) #obl tary
        click_kb(self.keyboard_controller, keyboard.Key.right, 2)
        click_kb(self.keyboard_controller, keyboard.Key.down) #ogr asort
        time.sleep(self.interval)
        click_kb(self.keyboard_controller, keyboard.Key.right)#reg zaokr
        click_kb(self.keyboard_controller, keyboard.Key.down) 
        time.sleep(self.interval)
        click_kb(self.keyboard_controller, keyboard.Key.right)
        click_kb(self.keyboard_controller, keyboard.Key.down) #akt opak zwrotu
        time.sleep(self.interval)
        click_kb(self.keyboard_controller, keyboard.Key.right, 2)
        click_kb(self.keyboard_controller, keyboard.Key.down) #wyznaczanie kursu
        time.sleep(self.interval)
        click_kb(self.keyboard_controller, keyboard.Key.right) 
        click_kb(self.keyboard_controller, keyboard.Key.down, 2) #obliczanie ubytku
        time.sleep(self.interval)
        click_kb(self.keyboard_controller, keyboard.Key.right)
        click_kb(self.keyboard_controller, keyboard.Key.down) #podzial
        click_kb(self.keyboard_controller, keyboard.Key.right, 2)
        time.sleep(self.interval)
        # click_kb(self.keyboard_controller, keyboard.Key.down)
        # click_kb(self.keyboard_controller, keyboard.Key.right, 2)
    
    def insert_price(self):
        self.mouse_controller.position = self.price_coords
        self.mouse_controller.click(mouse.Button.left)
        time.sleep(self.interval)

        click_kb(self.keyboard_controller, keyboard.Key.down) #wyzn warunku
        click_kb(self.keyboard_controller, keyboard.Key.right) 
        time.sleep(self.interval)
        click_kb(self.keyboard_controller, keyboard.Key.down) #wyzn cen
        click_kb(self.keyboard_controller, keyboard.Key.right) 
        time.sleep(self.interval)
        click_kb(self.keyboard_controller, keyboard.Key.down, 11) #cena główna
        self.keyboard_controller.type('11')
        time.sleep(self.interval)
        click_kb(self.keyboard_controller, keyboard.Key.down, 4) #stała cena zmaow
        click_kb(self.keyboard_controller, keyboard.Key.right) 
        time.sleep(self.interval)
        self.mouse_controller.position = self.const_price_coords
        self.mouse_controller.click(mouse.Button.left)

        time.sleep(self.interval)
        click_kb(self.keyboard_controller, keyboard.Key.enter)

    def insert_payrules(self):
        move_click(self.mouse_controller, self.pay_rules_coords)
        time.sleep(self.interval)
        move_click(self.mouse_controller, self.pay_rules_set_coords)
        time.sleep(self.interval)
        #click_kb(self.keyboard_controller, keyboard.Key.right)
        click_kb(self.keyboard_controller, keyboard.Key.enter)
        time.sleep(self.interval)
    
    def insert_docs(self):
        move_click(self.mouse_controller, self.docs_coords)
        time.sleep(self.interval)
        click_kb(self.keyboard_controller, keyboard.Key.down, 4) #doc zamow
        self.keyboard_controller.type('1')
        click_kb(self.keyboard_controller, keyboard.Key.enter)
        time.sleep(self.interval)
    
    def insert_etyciet(self, lang):
        lang_num = self.return_lang_num(lang)
        move_click(self.mouse_controller, self.etyciet_coords)
        time.sleep(self.interval)
        click_kb(self.keyboard_controller, keyboard.Key.down) #wazenie etykieta
        time.sleep(self.interval)
        self.keyboard_controller.type(lang_num)
        time.sleep(self.interval)
        click_kb(self.keyboard_controller, keyboard.Key.enter)
        time.sleep(self.interval)
        click_kb(self.keyboard_controller, keyboard.Key.down) #pozycja etykieta
        self.keyboard_controller.type(lang_num)
        click_kb(self.keyboard_controller, keyboard.Key.enter)
        time.sleep(self.interval)

cl = Clicker()
# cl.insert_order(1093)
cl.read_sheet('dod_klienci.xlsx')
start = 5764
time.sleep(2)
cl.do_action(70, 126)
# cl.insert_order(7255, 'cz')
# cl.insert_ord()
# cl.insert_price()
# cl.insert_payrules(4918
# )4918