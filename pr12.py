from pynput import mouse #import Button, Controller
from pynput import keyboard
import time
import openpyxl
import pyperclip
from controller_lib import click_kb, move_click

class Clicker():
    def __init__(self , letter):
        self.adress_coords = (207, 197)
        self.data2_coords = (160, 302)
        self.prow_mag_coord = (187, 478)
        self.jedn_coords = (237, 304)
        self.dalsz_jedn_coords = (95, 424)
        self.new_jedn_coords = (333, 695)
        self.sell_coords = (106, 238)
        self.OK_coords = (651, 620)
        self.ext_data_coords = (300, 305)
        self.VAT_coords = (623, 343)
        self.sort_group_coords = (391, 306)
        self.origin_coords = (287, 237)
        self.durability_coords = (673, 308)
        self.prep_coords = (576, 598)

        self.mouse_controller = mouse.Controller()
        self.keyboard_controller = keyboard.Controller()
        self.interval = 0.2
        self.letter = letter

    def read_sheet(self, wb_filename):
        wb = openpyxl.load_workbook(wb_filename)
        self.sheet = wb['swieze']
    
    def prep_name(self, name):
        words = [word for word in name.split()]
        line1 = ''
        line2 = ''
        add_to_line1 = True
        add_to_line2 = True

        for word in words:
            if len(line1 + word) < 50 and add_to_line1 == True:
                line1 += word + ' '
            elif len(line2 + word) < 50 and add_to_line2 == True:
                add_to_line1 = False
                line2 += word + ' '
        
        return[line1, line2]
    
    def pick_number(self, row):
        field = self.sheet['N' + str(row)].value
        if field == 'filet':
            return '120'
        elif field == 'udziec':
            return '130'
        elif field == 'podudzie':
            return '140'
        elif field == 'skrzydło':
            return '150'
        elif field == 'szyja':
            return '160'
        elif field == 'korpus':
            return '190'
        elif field == 'kości skóry':
            return '200'
        elif field == 'mięso drobne':
            return '170'
        elif field == "rozdrobnione":
            return '180'
        elif field == "tuszka":
            return '110'
        else:
            return '210'

    def insert_st(self):
        time.sleep(self.interval)
        move_click(self.mouse_controller, self.new_jedn_coords)
        time.sleep(2 * self.interval)

        self.keyboard_controller.type("st" + '\n')
        time.sleep(2*self.interval)
        move_click(self.mouse_controller, self.sell_coords)

        click_kb(self.keyboard_controller, keyboard.Key.tab, 2)
        time.sleep(2*self.interval)
        self.keyboard_controller.type("1" + '\n')
        time.sleep(2*self.interval)
        move_click(self.mouse_controller, self.OK_coords)

    def insert_pa(self):
        time.sleep(self.interval)
        move_click(self.mouse_controller, self.new_jedn_coords)
        time.sleep(self.interval)
        self.keyboard_controller.type("PA" + '\n')
        time.sleep(self.interval)
        move_click(self.mouse_controller, self.sell_coords)
        click_kb(self.keyboard_controller, keyboard.Key.tab)
        self.keyboard_controller.type("800" + '\n')
        move_click(self.mouse_controller, self.OK_coords)

    #POPRAWIĆ RĘCZNIE 58250

    def insert_data1(self, row, id):
        time.sleep(2*self.interval)
        name1 = self.sheet[self.letter + str(row)].value
        name2 = self.sheet['B' + str(row)].value

        # click_kb(self.keyboard_controller, keyboard.Key.tab)
        self.keyboard_controller.type('\t')
        time.sleep(self.interval)
        lines = self.prep_name(name1)
        self.keyboard_controller.type(lines[0] + '\t')
        self.keyboard_controller.type(lines[1] + '\t')
        time.sleep(self.interval)
        self.keyboard_controller.type(name2 + '\t')
        self.keyboard_controller.type(name2 + '\t')
        time.sleep(self.interval)
        self.keyboard_controller.type('904' + '\t')
        self.keyboard_controller.type(self.pick_number(row) + '\t')
        time.sleep(self.interval)

    def pick_shipment(self, id):
        return '260' if (int(id) - 90) % 100 == 0 else '160'
    
    def pick_storage(self, id):
        if (int(id) - 50) % 100 == 0:
            return '1'
        if (int(id) - 51) % 100 == 0:
            return '2'
        else:
            return '3'
    
    def insert_data2(self, id):
        time.sleep(self.interval)
        move_click(self.mouse_controller, self.data2_coords)
        time.sleep(self.interval)
        click_kb(self.keyboard_controller, keyboard.Key.tab)
        click_kb(self.keyboard_controller, keyboard.Key.tab)
        self.keyboard_controller.type(self.pick_shipment(id) + '\t')
        self.keyboard_controller.type(self.pick_storage(id) + '\t')
        time.sleep(self.interval)
        move_click(self.mouse_controller, self.prow_mag_coord)
        time.sleep(self.interval)
    
    def type0012(self):
        click_kb(self.keyboard_controller, '0')
        click_kb(self.keyboard_controller, '.')
        click_kb(self.keyboard_controller, '0')
        click_kb(self.keyboard_controller, '1')
        click_kb(self.keyboard_controller, '2')

    def insert_jedn(self, id):
        time.sleep(self.interval)
        move_click(self.mouse_controller, self.jedn_coords)
        time.sleep(self.interval)
        click_kb(self.keyboard_controller, keyboard.Key.tab)
        self.keyboard_controller.type('KG' + '\t')
        self.keyboard_controller.type('PJ' + '\t')
        self.keyboard_controller.type('20' + '\t')
        self.keyboard_controller.type('PJ' + '\t')
        self.keyboard_controller.type('KG' + '\t')
        self.keyboard_controller.type('PJ' + '\t')
        self.keyboard_controller.type('20' + '\t')
        if (int(id) - 50) % 100 != 0:
            time.sleep(10*self.interval)
            click_kb(self.keyboard_controller, keyboard.Key.tab, 2)
            time.sleep(2*self.interval)
            self.type0012()
            time.sleep(2*self.interval)

        time.sleep(self.interval)
        move_click(self.mouse_controller, self.dalsz_jedn_coords)
        time.sleep(self.interval)
        self.insert_st()
        if (id - 90) % 100 >= 0 and (id - 90) % 100 < 4:
            self.insert_pa()
        time.sleep(self.interval)

    def insert_ext_data(self):
        time.sleep(self.interval)
        move_click(self.mouse_controller, self.ext_data_coords)
        time.sleep(self.interval)
        move_click(self.mouse_controller, self.VAT_coords)
        time.sleep(2 * self.interval)
        self.keyboard_controller.type("1" + '\n')
        time.sleep(self.interval)
    
    def pick_sex(self, id):
        field = self.sheet['L' + str(i)].value
        if field == 'indor':
            return '1'
        elif field == 'indyczka':
            return '2'
        elif field == 'indyk':
            return '3'

    def pick_packaging(self, id):
        field = self.sheet['M' + str(i)].value
        if field == 'poliblok':
            return '3'
        elif field == 'luz':
            return '1'
        else:
            return '5'

    def insert_sort_group(self, id, row):
        time.sleep(self.interval)
        move_click(self.mouse_controller, self.sort_group_coords)
        time.sleep(self.interval)
        move_click(self.mouse_controller, self.prep_coords)
        time.sleep(self.interval)
        move_click(self.mouse_controller, self.origin_coords)
        time.sleep(self.interval)
        click_kb(self.keyboard_controller, keyboard.Key.f4)
        self.keyboard_controller.type('2\n\n')
        time.sleep(self.interval)
        click_kb(self.keyboard_controller, keyboard.Key.down, 2)
        click_kb(self.keyboard_controller, keyboard.Key.f4)
        self.keyboard_controller.type(self.pick_sex(row) + '\n\n')
        time.sleep(self.interval)
        click_kb(self.keyboard_controller, keyboard.Key.down, 2)
        click_kb(self.keyboard_controller, keyboard.Key.f4)
        self.keyboard_controller.type(self.pick_packaging(row) + '\n\n')
        if (id - 51) % 100 != 0:    
            time.sleep(self.interval)
            click_kb(self.keyboard_controller, keyboard.Key.down, 3)
            click_kb(self.keyboard_controller, keyboard.Key.f4)
            click_kb(self.keyboard_controller, keyboard.Key.down, 2)
            self.keyboard_controller.type('1\n\n')
        time.sleep(self.interval)
        click_kb(self.keyboard_controller, keyboard.Key.esc)

    def pick_durability(self, id):
        if (int(id) - 50) % 100 == 0 or (int(id) - 90) % 100 == 0:
            return '365'
        elif (int(id) - 51) % 100 == 0:
            return '6'
        elif (int(id) - 91) % 100 == 0:
            return '14'
        else:
            return '12'
    
    def insert_durability(self, id):
        time.sleep(self.interval)
        move_click(self.mouse_controller, self.durability_coords)
        time.sleep(self.interval)
        click_kb(self.keyboard_controller, keyboard.Key.down)
        time.sleep(self.interval)
        self.keyboard_controller.type(self.pick_durability(id) + '\n')

    def do_action(self, start, stop):
        for i in range(start, stop):
            id = self.sheet['A' + str(i)].value
            print(self.pick_durability(id))
            move_click(self.mouse_controller, self.adress_coords)
            time.sleep(self.interval)
            self.keyboard_controller.type(str(id) + "030")
            click_kb(self.keyboard_controller, keyboard.Key.enter)
            time.sleep(self.interval)
            self.insert_data1(i, id)
            self.insert_data2(id)
            self.insert_jedn(id)
            self.insert_ext_data()
            # self.insert_sort_group(id, i)
            self.insert_durability(id)
            time.sleep(2*self.interval)
            click_kb(self.keyboard_controller, keyboard.Key.f2)
            time.sleep(2*self.interval)
            click_kb(self.keyboard_controller, keyboard.Key.esc)
            print("Zrobiono klienta rzad: " + str(i) + " numer: " + str(id) + "030")
            time.sleep(2)

    
cl = Clicker('E')
cl.read_sheet('asort.xlsx')
# time.sleep(2)
# cl.insert_data1(13, 11150)
# cl.insert_ext_data()
# cl.insert_st()

cl.do_action(8, 12)

#483 614