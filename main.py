import sys
import time
import tkinter as tk
from tkinter import ttk
import os
from _tkinter import TclError
import csv
import threading
import ctypes
import pickle
import random as rd

import keyboard as kb
from PIL import Image, ImageTk
import pyautogui
import win32con as wcon
from pynput.keyboard import Controller, Listener
from pynput.mouse import Button, Controller as MouseController, Listener as MouseListener

__version__ = "0.1.0"


def unlock_keyboard():
    """Unlock keyboard"""
    kb.unhook_all()


def lock_keyboard():
    """Lock keyboard"""
    blocked_keys = []

    # Block other keys
    blocked_keys += ['caps lock', 'tab', 'windows', 'left arrow', 'right arrow'] + [f'f{i}' for i in range(1, 13)]
    blocked_keys += ['volume up', 'volume down', 'space',
                     'up arrow', 'down arrow', 'left arrow', 'right arrow',
                     'insert', 'home', 'page up', 'end', 'page down',
                     'pause', 'scroll lock', 'print screen']
    blocked_keys += list('abcdefghijklmnopqrstuvwxyz')

    for keyname in blocked_keys:
        kb.block_key(keyname)


DEBUG = False
if os.path.isfile('DEBUG'):
    DEBUG = True


class Main:
    def __init__(self, password):
        # Define
        self.root = tk.Tk()
        self.screen_size = {'width': self.root.winfo_screenwidth(), 'height': self.root.winfo_screenheight()}
        self.start_time = time.time()
        self.prev_position = pyautogui.position()
        self.password_start_time = self.start_time
        self.password = password
        self.window_is_sleeping = False
        self.lock = False
        self.class_table = []
        self.is_down = False
        self.already = []

        with open('data\\classTable.csv', newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile)
            for row in reader:
                self.class_table.append(row)
        self.tabel_labels = []
        try:
            tabel = self.class_table[int(time.strftime("%w"))]
        except IndexError:
            tabel = list("今天不上课！")
        print(time.strftime("%w"))
        for i, _class in enumerate(tabel):
            label = ttk.Label(self.root, text=_class[0], font=('Times', 32))
            self.tabel_labels.append(label)
            self.tabel_labels[i].place(x=2, y=i * self.tabel_labels[i].winfo_reqheight() + i)

        self.root.title("电教助手")
        self.root.attributes("-alpha", 0.8)  # Make window semi-transparent
        self.root.iconbitmap('.\\images\\icon.ico')
        self.root.overrideredirect(True)  # Remove title bar
        self.root.wm_attributes("-topmost", True)  # Keep window on top
        self.root.geometry(f"400x600+{self.screen_size['width'] - 405}+{self.screen_size['height'] // 2 - 300}")
        self.root.bind("<Button-1>", self.active_window)
        self.root.bind("<Button-2>", self.active_window)
        self._image = (ImageTk.PhotoImage(Image.open('images\\quit.png').resize((25, 25))),
                       ImageTk.PhotoImage(Image.open('images\\settings.png').resize((25, 25))))
        tk.Button(self.root, image=self._image[0], command=self.exit).place(x=365, y=1)
        tk.Button(self.root, image=self._image[1], command=self.settings).place(x=325, y=1)
        self.time = tk.Frame(self.root)
        self.time.place(x=90, y=35)
        self.time_label1 = tk.Label(self.time, text=time.strftime("%H:%M:%S"), font=('Arial', 40))
        self.time_label2 = tk.Label(self.time, text=time.strftime("%Y-%m-%d"), font=('Arial', 15))
        self.time_label1.pack()
        self.time_label2.pack()
        self.random = tk.LabelFrame(self.root, text='随机学号生成', font=('Arial', 20), width=300, height=350)
        self.random.place(x=90, y=180)
        self.choice = tk.BooleanVar(value=False)
        ttk.Checkbutton(self.random, text='去重模式', variable=self.choice).place(x=10, y=20)
        ttk.Label(self.random, text='范围：').place(x=0, y=0)
        self.start_entry = ttk.Entry(self.random, width=5)
        self.start_entry.place(x=40, y=0)
        tk.Label(self.random, text=' ~ ').place(x=95, y=0)
        self.end_entry = ttk.Entry(self.random, width=5)
        self.end_entry.place(x=120, y=0)
        self.random_label = tk.Label(self.random, text='04', font=('Times', 80, 'bold'), bg='white', fg='red',
                                     borderwidth=2, relief="ridge")
        self.random_label.place(x=70, y=40)
        tk.Button(self.random, text='生 成 随 机 学 号', font=('Times', 20, 'bold'), bg='red', fg='blue',
                  highlightthickness=4, command=self.generate_random).place(x=10, y=220)

        # Create keyboard controller
        self.keyboard = Controller()
        # Create mouse controller
        self.mouse = MouseController()

        # Create keyboard listener
        self.keyboard_listener = Listener(on_press=self.reset, on_release=self.reset)

        # Start listener threads
        threading.Thread(target=self.keyboard_listener.start, daemon=True).start()

    def generate_random(self):
        from tkinter import messagebox
        if not self.start_entry.get():
            self.start_entry.insert(0, '1')
        if not self.end_entry.get():
            self.end_entry.insert(0, '40')
        try:
            start, end = int(self.start_entry.get()), int(self.end_entry.get())
        except ValueError:
            messagebox.showeinfo("电教助手错误", "请输入0~99整数随机范围!")
            return
        if start < 0:
            start = abs(start)
        if end < 0:
            end = abs(end)
        if start > end:
            start, end = end, start
        can_choice = list(range(start, end + 1))
        for i in reversed(range(start - 1, end)):
            if can_choice[i] in self.already:
                can_choice.pop(i)
        if len(can_choice) == 0:
            messagebox.showinfo("电教助手随机学号", "学号已抽完!")
            self.already = []
            return
        if len(can_choice) == 1:
            num = can_choice[0]
            self.already.append(num)
        elif self.choice.get():
            num = rd.choice(can_choice)
            self.already.append(num)
        else:
            num = rd.randint(start, end)
        if len(str(num)) == 1:
            self.random_label['text'] = '0' + str(num)
            return
        self.random_label['text'] = str(num)

    def reset(self, *args, **kwargs):
        self.password_start_time = time.time()

    def exit(self):
        self.keyboard_listener.stop()
        self.root.destroy()
        sys.exit()

    def run(self):
        global DEBUG
        while True:
            self.time_label1['text'] = time.strftime("%H:%M:%S")
            self.time_label2['text'] = time.strftime("%Y-%m-%d")
            if self.prev_position != pyautogui.position():
                self.reset()
                self.prev_position = pyautogui.position()
            elif bool(ctypes.windll.user32.GetAsyncKeyState(0x01) & 0x8000) or bool(
                    ctypes.windll.user32.GetAsyncKeyState(0x02) & 0x8000):
                self.reset()
            self.root.update()
            _ = 3 if DEBUG else 25
            if time.time() - self.start_time >= _ and not self.window_is_sleeping:
                self.root.geometry(f"60x600+{self.screen_size['width'] - 65}+{self.screen_size['height'] // 2 - 300}")
                self.window_is_sleeping = True
                self.root.wm_attributes('-alpha', 0.5)
            _ = 10 if DEBUG else 500
            if time.time() - self.password_start_time >= _ and not self.lock:
                self.start_lock()

    def active_window(self, _=''):
        self.start_time = time.time()
        if not self.window_is_sleeping:
            return
        self.root.geometry(f"400x600+{self.screen_size['width'] - 405}+{self.screen_size['height'] // 2 - 300}")
        self.window_is_sleeping = False
        self.root.wm_attributes('-alpha', 0.8)

    def settings(self):
        from tkinter import simpledialog, messagebox
        password = simpledialog.askstring("设置", "请输入锁屏解锁密码", show='*')
        if password == self.password:
            password1 = ''
            password2 = ' '
            while password1 != password2:
                password1 = simpledialog.askinteger("设置", "请输入新的锁屏解锁密码")
                if not password1:
                    messagebox.showinfo("设置", "密码不能为空！")
                    continue
                password2 = simpledialog.askinteger("设置", "请再次确认新的锁屏解锁密码")
                if not password2:
                    messagebox.showinfo("设置", "密码不能为空！")
                    continue
                elif password1 != password2:
                    messagebox.showinfo("设置", "两次输入的密码不一致！")
                    continue
                messagebox.showinfo("设置", "密码设置成功！")
                with open('data\\data.dat', 'wb') as fb:
                    pickle.dump(int(password1), fb)
                self.password = str(password1)
        else:
            messagebox.showinfo("设置", "密码错误！")

    def start_lock(self):
        self.lock = True
        lock_keyboard()
        is_running_password = False
        window = tk.Tk()

        def show_password_window():
            nonlocal is_running_password
            if is_running_password:
                return
            is_running_password = True
            from tkinter import simpledialog, messagebox
            password_window = tk.Tk()
            password_window.title("解锁")
            password_window.wm_attributes('-topmost', True)
            password_window.iconbitmap('.\\images\\icon.ico')
            entry = ttk.Entry(password_window, show='*')
            entry.pack()

            def try_password(*args, **kwargs):
                nonlocal entry, password_window, is_running_password, window, self
                password = entry.get()
                if password == self.password:
                    self.lock = False
                    password_window.destroy()
                    window.destroy()
                    unlock_keyboard()
                    self.password_start_time = time.time()
                    is_running_password = False
                    return
                messagebox.showinfo("解锁", "密码错误！")
                is_running_password = False

            ttk.Button(password_window, text='确定', command=try_password).pack()
            password_window.mainloop()

        window.wm_attributes("-topmost", True)  # Keep window always on top
        window.wm_attributes("-fullscreen", True)  # Make window full screen
        window.wm_attributes("-alpha", 0.1)
        window.title("电教助手锁屏")
        window.iconbitmap('.\\images\\icon.ico')
        window.overrideredirect(True)  # Remove title bar
        button = tk.Button(window, text='', font=('Times', 20), width=200, height=200, command=show_password_window)
        button.pack()
        window.mainloop()


if __name__ == "__main__":
    with open('.\\data\\data.dat', 'rb') as f:
        psd = str(pickle.load(f))
    try:
        Main(psd).run()
    except TclError:
        pass
