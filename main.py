from pynput import mouse  # import Button, Controller
from pynput import keyboard
import time
import openpyxl
import pyperclip
from controller_lib import Base_Clicker
PODROBY_ID = [81291, 81391, 81491, 81591,
              84291, 84391, 84491, 87291, 87391, 87491]
KLB_ID = [21191, 31191, 34191, 41191, 44191, 59190, 59191, 72090]


class Clicker(Base_Clicker):
    def __init__(self, wb_filename, template_fr_filename, template_other_filename, template_info_filename, dates_fr_filename, dates_ot_filename):
        super().__init__()
        self.addr_cords = (242, 355)
        self.interval = 0.4
        with open("txt_files/" + template_fr_filename, encoding="utf8") as f:
            self.template_lines_fr = f.read().splitlines()
        with open("txt_files/" + template_other_filename, encoding="utf8") as f:
            self.template_lines_other = f.read().splitlines()
        with open("txt_files/" + template_info_filename, encoding="utf8") as f:
            self.template_lines_info = f.read().splitlines()
        with open("txt_files/" + dates_fr_filename, encoding="utf8") as f:
            self.dates_fr = f.read().splitlines()
        with open("txt_files/" + dates_ot_filename, encoding="utf8") as f:
            self.dates_ot = f.read().splitlines()

    def do_action(self, column_letter, split_word, start, stop):
        for i in range(start, stop):  # self.sheet.max_row
            id = self.sheet['A' + str(i)].value
            if((int(id) - 90) % 100 < 4):
                self.insert_vacuum(i, id, column_letter, split_word)
            else:
                self.insert_other(i, id, column_letter)
            with open('log.txt', 'a') as f:
                message = 'Zapisano ' + \
                    str(i) + ' ' + str(id) + ' progres' + \
                    str(i*100/self.sheet.max_row) + '%'
                f.write(message)
                print(message)
            time.sleep(0.5)

    def insert_vacuum(self, index, id, column_letter, split_word):
        lines = self.shorten_name(
            self.sheet[column_letter + str(index)].value.replace("KLASSE B", ''), split_word)
        class_letter = "B" if int(id) in KLB_ID else "A"
        additional1 = "HKL " + class_letter + ", VACUUMIERT"
        self.mouse_controller.position = (self.addr_cords)
        self.mouse_controller.click(mouse.Button.left, 1)
        time.sleep(self.interval)
        self.keyboard_controller.type(str(id))
        self.click_kb(keyboard.Key.enter)
        time.sleep(self.interval)
        for line in lines:
            if line == "":
                self.keyboard_controller.type('\b' + '\t')
            else:
                self.keyboard_controller.type(line + '\t')

        self.keyboard_controller.type(additional1 + '\t')
        self.keyboard_controller.type('\b' + '\t')

        time.sleep(self.interval)
        self.click_kb(keyboard.Key.f10, 2)
        time.sleep(self.interval)
        if int(id) % 10 == 0:
            self.insert_frozen()
        else:
            self.insert_fresh(id)

        time.sleep(self.interval)
        self.click_kb(keyboard.Key.f10)

        time.sleep(self.interval)
        for line in self.template_lines_info:
            self.keyboard_controller.type(line + '\t')

        time.sleep(self.interval)
        self.click_kb(keyboard.Key.f2)

    def insert_frozen(self):
        temperature_txt = "Legerung bei: -18°C"
        for i, line in enumerate(self.template_lines_fr):
            if i == 1:
                pyperclip.copy(temperature_txt)
                time.sleep(self.interval)
                with self.keyboard_controller.pressed(keyboard.Key.ctrl):
                    self.click_kb('v')
            else:
                self.keyboard_controller.type(line)
            self.keyboard_controller.type('\t')
        self.click_kb(keyboard.Key.tab, 2)
        for line in self.dates_fr:
            self.keyboard_controller.type(line + '\t')

    def insert_fresh(self, id):
        number = 3 if int(id) in PODROBY_ID else 4
        temperature_txt = f"Bei sachgemäßer lagerung von 0°C bis +{number}°C:"
        for i, line in enumerate(self.template_lines_other):
            if i == 1:
                pyperclip.copy(temperature_txt)
                time.sleep(self.interval)
                with self.keyboard_controller.pressed(keyboard.Key.ctrl):
                    self.keyboard_controller.press('v')
                    self.keyboard_controller.release('v')
            elif i == 2:
                self.keyboard_controller.type("\b")
            else:
                self.keyboard_controller.type(line)
            self.keyboard_controller.type('\t')
        self.click_kb(keyboard.Key.tab, 2)
        for line in self.dates_ot:
            self.keyboard_controller.type(line + '\t')

    def insert_other(self, index, id, column_letter):
        lines = self.shorten_name(
            self.sheet[column_letter + str(index)].value, "żżż")

        self.mouse_controller.position = (self.addr_cords)
        self.mouse_controller.click(mouse.Button.left, 1)
        time.sleep(self.interval)
        self.keyboard_controller.type(str(id))
        self.click_kb(keyboard.Key.enter)
        time.sleep(self.interval)
        for line in lines:
            self.keyboard_controller.type(line + '\t')
        time.sleep(self.interval)
        self.click_kb(keyboard.Key.f2)

    def shorten_name(self, name_string, split_word):
        name_string = name_string.replace(split_word, '')
        words = [word for word in name_string.split()]
        line1 = ''
        line2 = ''
        line3 = ''

        add_to_line1 = True
        add_to_line2 = True

        for word in words:
            if len(line1 + word) < 40 and add_to_line1 == True:
                line1 += word + ' '
            elif len(line2 + word) < 40 and add_to_line2 == True:
                add_to_line1 = False
                line2 += word + ' '
            else:
                add_to_line2 = False
                line3 += word + ' '

        line1 = line1.strip()
        line2 = line2.strip()
        line3 = line3.strip()

        return [line1, line2, line3]


c = Clicker("sheets/asort.xlsx", 'frozen.txt', "rest.txt",
            "other_info.txt", "dates_fr.txt", "dates_ot.txt")
time.sleep(2)
c.read_sheet("asort.xls", "swieze")
# c.insert_other(8, 11050, 'F')
# c.insert_other(9, 11051, 'F')
c.do_action('E', 'VACUUM', 140, 695)
# c.insert_vacuum(302, 34191, 'E', 'VIDE')
#print('MANTENER EN LA TEMPERATURA DE 0 ° C - + 4 ° C')
# 0-14291 POPRAWIĆ
