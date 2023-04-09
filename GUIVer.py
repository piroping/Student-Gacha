import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox
import random
import datetime

import openpyxl

class Screen(ttk.Frame):
    def __init__(self, master):
        ttk.Frame.__init__(self,master=master)
        
        self.output_name = tk.StringVar()
        self.output_name.set('ガチャ')
        
        self.label_output_name = ttk.Label(textvariable=self.output_name, anchor=tk.CENTER,
                                           font=('MSゴシック', 80, 'bold'))
        
        self.button_start_lottery = ttk.Button(text='Press Me', command=read_file.output_name_random)
        self.button_setting_by_time = ttk.Button(text='時間帯', command=self.change_by_time)
        self.selecting = tk.StringVar()
        self.combobox_sheet_selecting = ttk.Combobox(state='readonly', values=read_file.sheet_names,
                                                     textvariable=self.selecting)
        self.combobox_sheet_selecting.bind('<<ComboboxSelected>>', self.check_can_use)
        
        
        self.label_output_name.grid(column=0, row=0, columnspan=3, sticky=tk.W + tk.E)
        self.button_start_lottery.grid(column=0, row=1, columnspan=3, sticky=tk.W + tk.E)
        self.button_setting_by_time.grid(column=0, row=2, sticky=tk.W + tk.E)
        self.combobox_sheet_selecting.grid(column=1, row=2, sticky=tk.W + tk.E)

        
        self.check_can_use(None)
        self.change_by_time()
    
    def check_can_use(self, _):
        if self.selecting.get() == '':
            self.button_start_lottery['state'] = 'disable'
        else:
            self.button_start_lottery['state'] = 'normal'
    
    def change_by_time(self):
        if (c := read_file.time_set()) is not None:
            self.selecting.set(c)

class Read_Xlsx_File():
    def __init__(self):
        try:
            self.wb = openpyxl.load_workbook('Student Gacha.xlsx')
            self.sheet_names = self.wb.sheetnames
            self.sheet_names.remove('設定')
        except FileNotFoundError:
            messagebox.showerror('Error', 'Student Gacha.xlsxが見つかりませんでした')
            exit()

    def output_name_random(self):
        if screen.selecting.get() == '':
            return None
        self.target = list(self.wb[screen.selecting.get()].values)
        try:
            self.colors = self.target[0][1:]
        except IndexError:
            messagebox.showerror('Error', '選択されたページに何も書かれていません')
            return
        self.members = {i:[] for i in self.colors}
        self.probabilities = [int(i) for i in self.target[1][1:]]
        for i in self.target[2:]:
            if i[0] is None or i[0] == '':
                break
            for x, y in zip(i[1:], self.colors):
                if x is not None:
                    self.members[y].append(i[0])
                    break
        while True:
            try:
                calc = random.randint(1, sum(self.probabilities))
            except ValueError:
                messagebox.showerror('Error', '選択されたページに何も書かれていません')
                return
            for x, y in zip(self.probabilities, self.colors):
                calc -= x
                if calc <= 0:
                    try:
                        c = random.choice(self.members[y])
                    except IndexError:
                        continue
                    if screen.output_name.get() == c:
                        continue
                    screen.output_name.set(c)
                    break
            else:
                continue
            break
    
    def time_set(self):
        dt = datetime.datetime.now()
        if dt.isoweekday() == 7:
            return
        calc = dt - datetime.timedelta(hours=8, minutes=30) # 一時間目スタート
        frames = calc.hour * 60 + calc.minute
        if frames >= 270:
            frames += 5 # 昼休み
        frames //= 55
        if 0 <= frames <= 7:
            out = list(self.wb['設定'].values)[frames + 1][dt.isoweekday()]
            if out in self.sheet_names:
                return out

if __name__ == '__main__':
    read_file = Read_Xlsx_File()
    
    root = tk.Tk()
    root.geometry('750x250')
    root.title('School Gacha')
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    screen = Screen(root)
    screen.mainloop()