from pynput import mouse #import Button, Controller
from pynput import keyboard
import time
import openpyxl
import pyperclip

class Clicker():
    def __init__(self, wb_filename):
        self.sheet = self.read_sheet(wb_filename)
        self.mouse_controller = mouse.Controller()
        self.keyboard_controller = keyboard.Controller()
        self.cursor_x = 206
        self.cursor_y = 358
        self.interval = 0.2

    def read_sheet(self, wb_filename):
        wb = openpyxl.load_workbook(wb_filename)
        sheet = wb['swieze']
        return sheet
    
    def do_action(self, column_letter, vacuum, start, stop=12):
        additional1 = "CLASSE A, VACUUM"
        for i in range(start, stop): #self.sheet.max_row
            id = self.sheet['A' + str(i)].value
            if (int(id) - 90) % 100 > 3:
                continue
            
            lines = self.shorten_name(self.sheet[column_letter + str(i)].value, vacuum)
        
            self.mouse_controller.position = (self.cursor_x, self.cursor_y)
            self.mouse_controller.click(mouse.Button.left, 1)
            time.sleep(self.interval)
            self.keyboard_controller.type(str(id))
            self.keyboard_controller.press(keyboard.Key.enter)
            self.keyboard_controller.release(keyboard.Key.enter)
            time.sleep(self.interval)
            for line in lines:
                if line == "":
                    self.keyboard_controller.type('\b' + '\t')
                else:
                    self.keyboard_controller.type(line + '\t')

            if (int(id) - 90) % 100 < 4:
                self.keyboard_controller.type(additional1 + '\t')
                self.keyboard_controller.type('\b' + '\t')
            time.sleep(self.interval)
            self.keyboard_controller.press(keyboard.Key.f2)
            self.keyboard_controller.release(keyboard.Key.f2)
            with open('log.txt', 'a') as f:
                f.write('Zmieniono ' + str(i) + ' ' + str(id) + '\n')
            time.sleep(self.interval)
    
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
        if len(line1) > 40 or len(line2) > 40 or len(line3) > 40:
            print('xd')
        print([line1, line2, line3])
        return [line1, line2, line3]

time.sleep(2)
cl = Clicker("asort.xlsx")
cl.do_action('I', 'VIDE', 8, 695)