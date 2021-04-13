from pynput import mouse #import Button, Controller
from pynput import keyboard
import time

mc = mouse.Controller()
time.sleep(1)
mc.position = (571, 43)
time.sleep(0.1)
mc.press(mouse.Button.left)
time.sleep(0.1)
mc.position = (100, 51)
time.sleep(0.1)
mc.release(mouse.Button.left)
