import json
import os.path
import re
import tkinter as tk
from tkinter import ttk, filedialog
import numpy as np
import pandas as pd
from typing import Union, TypeAlias, Self
import time
import win32com.client
import pythoncom
import math
from fractions import Fraction


RowColumnNum: TypeAlias = tuple[int, int]
Vector: TypeAlias = list[Union[int, float]]


def cell_ref2rc(reference: str) -> tuple:
    result = re.search('([A-Z]+)([0-9]+)', reference)
    col_title = result.group(1)
    row_num = int(result.group(2))
    col_num = 0
    for B in range(len(col_title)):
        col_num *= 26
        col_num += ord(col_title[B]) - ord('A') + 1
    return row_num, col_num


class ExcelReader:
    def __init__(self, dataframe: pd.DataFrame) -> None:
        self.df = dataframe

    def read_cell(self, reference: Union[str, RowColumnNum]):
        if isinstance(reference, str):
            row_num, col_num = cell_ref2rc(reference)
        else:
            row_num, col_num = reference
        return self.df.iloc[row_num - 1, col_num - 1]


class App(tk.Tk):
    def __init__(self, sheet_names_list: list) -> None:
        super().__init__()

        self.list = sheet_names_list

        self.title('Dibujar vigas')
        self.iconbitmap("panda-icon.ico")
        self.was_cancelled = True
        self.selected_indexes = None
        self.listbox = None
        self.geometry("250x300")
        self.minsize(250, 200)
        self.show_list()
        self.lift()
        self.mainloop()

    def show_list(self):
        top_frame = tk.Frame(self)
        sel_frame = tk.Frame(self)
        list_frame = tk.Frame(self)
        bottom_frame = tk.Frame(self)

        top_frame.pack(side=tk.TOP, padx=10, pady=2.5)
        sel_frame.pack(side=tk.TOP, padx=10, pady=2.5)
        list_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=2.5)
        bottom_frame.pack(side=tk.BOTTOM, padx=10, pady=2.5)

        tk.Label(self, text='Seleccione las hojas a dibujar:').pack(in_=top_frame)
        var = tk.Variable(value=self.list)
        self.listbox = tk.Listbox(
            self,
            listvariable=var,
            height=6,
            selectmode=tk.EXTENDED)
        select_all_button = ttk.Button(
            self,
            text='Select all',
            command=self.select_all)
        select_all_button.pack(in_=sel_frame, side=tk.LEFT, padx=10)
        deselect_all_button = ttk.Button(
            self,
            text='Deselect all',
            command=self.deselect_all)
        deselect_all_button.pack(in_=sel_frame, side=tk.LEFT, padx=10)
        self.listbox.pack(in_=list_frame, side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(
            list_frame,
            orient=tk.VERTICAL,
            command=self.listbox.yview
        )
        self.listbox['yscrollcommand'] = scrollbar.set
        scrollbar.pack(side=tk.LEFT, fill=tk.Y)
        ok_button = ttk.Button(
            self,
            text='OK',
            command=self.save_and_destroy)
        ok_button.pack(in_=bottom_frame, side=tk.LEFT, padx=10)
        cancel_button = ttk.Button(
            self,
            text='Cancel',
            command=self.destroy)
        cancel_button.pack(in_=bottom_frame, side=tk.LEFT, padx=10)

    def save_and_destroy(self):
        self.was_cancelled = False
        self.selected_indexes = self.listbox.curselection()
        self.destroy()

    def select_all(self):
        self.listbox.select_set(0, tk.END)

    def deselect_all(self):
        self.listbox.selection_clear(0, tk.END)


class AskFileApp(tk.Tk):
    def __init__(self, filetypes: tuple) -> None:
        super().__init__()
        self.withdraw()
        self.path = None
        self.filetypes = filetypes
        self.title('Dibujar vigas')
        self.iconbitmap("panda-icon.ico")
        self.was_cancelled = True

    def ask_file(self) -> None:
        self.path = filedialog.askopenfilename(filetypes=self.filetypes)
        if self.path:
            self.was_cancelled = False
            self.destroy()


class Assistant:
    def __init__(self, jsonFileName='beams_info', xlsxFilePath=None) -> None:
        self.fileName = jsonFileName
        if not xlsxFilePath:
            fileApp = AskFileApp(filetypes=(("Excel files", "*xlsx"), ("Excel files", "*xlsm")))
            fileApp.ask_file()
            xlsxFilePath = fileApp.path
            if fileApp.was_cancelled:
                exit()
        self.jsonFilePath = os.path.join(os.path.dirname(xlsxFilePath), jsonFileName + '.json')
        self.jsonFile = None
        self.jsonDict = dict()
        self.workbook = pd.ExcelFile(xlsxFilePath)

    def read_json_file(self):
        if not os.path.exists(self.jsonFilePath):
            raise Exception(f"No se existe el archivo: {self.jsonFilePath}")
        self.jsonFile = open(self.jsonFilePath, 'r')
        self.jsonDict = json.load(self.jsonFile)

    def create_json_file(self):
        self.jsonFile = open(self.jsonFilePath, 'w')
        json.dump(self.jsonDict, self.jsonFile)

    def get_sheets2draw(self) -> tuple:
        sheets_window = App(self.workbook.sheet_names[6:])
        sheets_window.mainloop()
        if sheets_window.was_cancelled:
            exit()
        else:
            return sheets_window.selected_indexes

    def download_excel_beams_info(self, indexes_list: Union[list, tuple, np.ndarray]):
        for index in indexes_list:
            span_info = self.download_excel_span_info(index)
            beam_name = re.findall('(.+?)\(', span_info['span_name'])[0]
            beam_dict = self.jsonDict.setdefault(beam_name,
                                                 {
                                                     "beam_name": beam_name,
                                                     "spans_num": 0,
                                                     "spans_info": []
                                                 })
            beam_dict['spans_num'] += 1
            beam_dict['spans_info'].append(span_info)

    def download_excel_span_info(self, index):
        sheet = ExcelReader(self.workbook.parse(index, header=None))
        span_keys = ['span_name', 'left_support_info', 'right_support_info', 'free_length', 'width', 'height',
                     'bars_info', 'stirrups_info']
        span_name = self.workbook.sheet_names[index]
        ls_info = [sheet.read_cell('L78'), sheet.read_cell('N78')]
        rs_info = [sheet.read_cell('L80'), sheet.read_cell('N80')]
        free_length = sheet.read_cell('L79')
        width = sheet.read_cell('Q78') / 100
        height = sheet.read_cell('Q79') / 100

        bars_info = {
            'annotated_dimensions': {
                'top_left': list(),
                'top_right': list(),
                'bottom_left': list(),
                'bottom_center': list(),
                'bottom_right': list()},
            'quantity': {
                'top_left': 0,
                'top_right': 0,
                'bottom_left': 0,
                'bottom_center': 0,
                'bottom_right': 0},
            'info': list()}
        bars_keys = ['label', 'case', 'side', 'order', 'left_cut', 'right_cut', 'tie_info']
        # Top long bar
        if sheet.read_cell('O90') != 0 or sheet.read_cell('O91') != 0:
            label = ""
            label_1 = str(int(sheet.read_cell('O90'))) + '%%C' + sheet.read_cell('Q90')
            label_2 = str(int(sheet.read_cell('O91'))) + '%%C' + sheet.read_cell('Q91')
            if sheet.read_cell('O90') != 0:
                label = label + label_1 + (' + ' if sheet.read_cell('O91') != 0 else '')
            if sheet.read_cell('O91') != 0:
                label = label + label_2
            case = 0
            side = 1
            order = 0
            left_cut = round(min(sheet.read_cell('F189') - sheet.read_cell('F219'),
                                 sheet.read_cell('F190') - sheet.read_cell('F219')), 2)
            right_cut = round(max(sheet.read_cell('F258') - sheet.read_cell('F257'),
                                  sheet.read_cell('F259') - sheet.read_cell('F257')), 2)
            tie_info = [[label_1, label_2], False, [False, False], False, [False, False]]
            tie_info[1] = True if sheet.read_cell('E97') else False
            tie_info[2][0] = True if sheet.read_cell('N90') else False
            tie_info[2][1] = True if sheet.read_cell('N91') else False
            tie_info[3] = True if sheet.read_cell('AA97') else False
            tie_info[4][0] = True if sheet.read_cell('R90') else False
            tie_info[4][1] = True if sheet.read_cell('R91') else False
            bars_info['info'].append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut, tie_info])))
        # Bottom long bar
        if sheet.read_cell('O135') != 0 or sheet.read_cell('O136') != 0:
            label = ""
            label_1 = str(int(sheet.read_cell('O135'))) + '%%C' + sheet.read_cell('Q135')
            label_2 = str(int(sheet.read_cell('O136'))) + '%%C' + sheet.read_cell('Q136')
            if sheet.read_cell('O135') != 0:
                label = label + label_1 + (' + ' if sheet.read_cell('O136') != 0 else '')
            if sheet.read_cell('O136') != 0:
                label = label + label_2
            case = 0
            side = -1
            order = 0
            left_cut = round(min(sheet.read_cell('F204') - sheet.read_cell('F219'),
                                 sheet.read_cell('F205') - sheet.read_cell('F219')), 2)
            right_cut = round(max(sheet.read_cell('F273') - sheet.read_cell('F257'),
                                  sheet.read_cell('F274') - sheet.read_cell('F257')), 2)
            tie_info = [[label_1, label_2], False, [False, False], False, [False, False]]
            tie_info[1] = True if sheet.read_cell('E125') else False
            tie_info[2][0] = True if sheet.read_cell('N135') else False
            tie_info[2][1] = True if sheet.read_cell('N136') else False
            tie_info[3] = True if sheet.read_cell('AA125') else False
            tie_info[4][0] = True if sheet.read_cell('R135') else False
            tie_info[4][1] = True if sheet.read_cell('R136') else False
            bars_info['info'].append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut, tie_info])))
        # Top left first order bar
        if sheet.read_cell('J101') != 0 or sheet.read_cell('J103') != 0:
            label = ""
            label_1 = str(int(sheet.read_cell('J101'))) + '%%C' + sheet.read_cell('L101')
            label_2 = str(int(sheet.read_cell('J103'))) + '%%C' + sheet.read_cell('L103')
            if sheet.read_cell('J101') != 0:
                label = label + label_1 + (' + ' if sheet.read_cell('J103') != 0 else '')
            if sheet.read_cell('J103') != 0:
                label = label + label_2
            case = 1
            side = 1
            order = 1
            left_cut = round(min(sheet.read_cell('F194') - sheet.read_cell('F219'),
                                 sheet.read_cell('F195') - sheet.read_cell('F219')), 2)
            right_cut = round(max(sheet.read_cell('F220') - sheet.read_cell('F219'),
                                  sheet.read_cell('F220') - sheet.read_cell('F219')), 2)
            tie_info = [[label_1, label_2], False, [False, False]]
            tie_info[1] = True if sheet.read_cell('E99') else False
            tie_info[2][0] = True if sheet.read_cell('I101') else False
            tie_info[2][1] = True if sheet.read_cell('I103') else False
            bars_info['quantity']['top_left'] += 1
            bars_info['annotated_dimensions']['top_left'].append(right_cut)
            bars_info['info'].append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut, tie_info])))
        # Top left second order bar
        if sheet.read_cell('F103') != 0 or sheet.read_cell('F105') != 0:
            label = ""
            label_1 = str(int(sheet.read_cell('F103'))) + '%%C' + sheet.read_cell('H103')
            label_2 = str(int(sheet.read_cell('F105'))) + '%%C' + sheet.read_cell('H105')
            if sheet.read_cell('F103') != 0:
                label = label + label_1 + (' + ' if sheet.read_cell('F105') != 0 else '')
            if sheet.read_cell('F105') != 0:
                label = label + label_2
            case = 1
            side = 1
            order = 2
            left_cut = round(min(sheet.read_cell('F199') - sheet.read_cell('F219'),
                                 sheet.read_cell('F200') - sheet.read_cell('F219')), 2)
            right_cut = round(max(sheet.read_cell('F223') - sheet.read_cell('F219'),
                                  sheet.read_cell('F223') - sheet.read_cell('F219')), 2)
            tie_info = [[label_1, label_2], False, [False, False]]
            tie_info[1] = True if sheet.read_cell('E101') else False
            tie_info[2][0] = True if sheet.read_cell('E103') else False
            tie_info[2][1] = True if sheet.read_cell('E105') else False
            bars_info['quantity']['top_left'] += 1
            bars_info['annotated_dimensions']['top_left'].append(right_cut)
            bars_info['info'].append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut, tie_info])))
        # Bottom left first order bar
        if sheet.read_cell('J118') != 0 or sheet.read_cell('J120') != 0:
            label = ""
            label_1 = str(int(sheet.read_cell('J118'))) + '%%C' + sheet.read_cell('L118')
            label_2 = str(int(sheet.read_cell('J120'))) + '%%C' + sheet.read_cell('L120')
            if sheet.read_cell('J118') != 0:
                label = label + label_1 + (' + ' if sheet.read_cell('J120') != 0 else '')
            if sheet.read_cell('J120') != 0:
                label = label + label_2
            case = 1
            side = -1
            order = 1
            left_cut = round(min(sheet.read_cell('F209') - sheet.read_cell('F219'),
                                 sheet.read_cell('F210') - sheet.read_cell('F219')), 2)
            right_cut = round(max(sheet.read_cell('F226') - sheet.read_cell('F219'),
                                  sheet.read_cell('F226') - sheet.read_cell('F219')), 2)
            tie_info = [[label_1, label_2], False, [False, False]]
            tie_info[1] = True if sheet.read_cell('E123') else False
            tie_info[2][0] = True if sheet.read_cell('I118') else False
            tie_info[2][1] = True if sheet.read_cell('I120') else False
            bars_info['quantity']['bottom_left'] += 1
            bars_info['annotated_dimensions']['bottom_left'].append(right_cut)
            bars_info['info'].append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut, tie_info])))
        # Bottom left second order bar
        if sheet.read_cell('F116') != 0 or sheet.read_cell('F118') != 0:
            label = ""
            label_1 = str(int(sheet.read_cell('F116'))) + '%%C' + sheet.read_cell('H116')
            label_2 = str(int(sheet.read_cell('F118'))) + '%%C' + sheet.read_cell('H118')
            if sheet.read_cell('F116') != 0:
                label = label + label_1 + (' + ' if sheet.read_cell('F118') != 0 else '')
            if sheet.read_cell('F118') != 0:
                label = label + label_2
            case = 1
            side = -1
            order = 2
            left_cut = round(min(sheet.read_cell('F214') - sheet.read_cell('F219'),
                                 sheet.read_cell('F215') - sheet.read_cell('F219')), 2)
            right_cut = round(max(sheet.read_cell('F229') - sheet.read_cell('F219'),
                                  sheet.read_cell('F229') - sheet.read_cell('F219')), 2)
            tie_info = [[label_1, label_2], False, [False, False]]
            tie_info[1] = True if sheet.read_cell('E121') else False
            tie_info[2][0] = True if sheet.read_cell('E116') else False
            tie_info[2][1] = True if sheet.read_cell('E118') else False
            bars_info['quantity']['bottom_left'] += 1
            bars_info['annotated_dimensions']['bottom_left'].append(right_cut)
            bars_info['info'].append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut, tie_info])))
        # Bottom central first order bar
        if sheet.read_cell('O116') != 0 or sheet.read_cell('O118') != 0:
            label = ""
            label_1 = str(int(sheet.read_cell('O116'))) + '%%C' + sheet.read_cell('Q116')
            label_2 = str(int(sheet.read_cell('O118'))) + '%%C' + sheet.read_cell('Q118')
            if sheet.read_cell('O116') != 0:
                label = label + label_1 + (' + ' if sheet.read_cell('O118') != 0 else '')
            if sheet.read_cell('O118') != 0:
                label = label + label_2
            case = 2
            side = -1
            order = 1
            left_cut = round(min(sheet.read_cell('F232') - sheet.read_cell('F219'),
                                 sheet.read_cell('F232') - sheet.read_cell('F219')), 2)
            right_cut = round(max(sheet.read_cell('F251') - sheet.read_cell('F257'),
                                  sheet.read_cell('F251') - sheet.read_cell('F257')), 2)
            tie_info = [[label_1, label_2], False, False]
            bars_info['quantity']['bottom_center'] += 1
            bars_info['annotated_dimensions']['bottom_center'].append((left_cut, right_cut))
            bars_info['info'].append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut, tie_info])))
        # Bottom central second order bar
        if sheet.read_cell('O108') != 0 or sheet.read_cell('O110') != 0:
            label = ""
            label_1 = str(int(sheet.read_cell('O108'))) + '%%C' + sheet.read_cell('Q108')
            label_2 = str(int(sheet.read_cell('O110'))) + '%%C' + sheet.read_cell('Q110')
            if sheet.read_cell('O108') != 0:
                label = label + label_1 + (' + ' if sheet.read_cell('O110') != 0 else '')
            if sheet.read_cell('O110') != 0:
                label = label + label_2
            case = 2
            side = -1
            order = 2
            left_cut = round(min(sheet.read_cell('F235') - sheet.read_cell('F219'),
                                 sheet.read_cell('F235') - sheet.read_cell('F219')), 2)
            right_cut = round(max(sheet.read_cell('F254') - sheet.read_cell('F257'),
                                  sheet.read_cell('F254') - sheet.read_cell('F257')), 2)
            tie_info = [[label_1, label_2], False, False]
            bars_info['quantity']['bottom_center'] += 1
            bars_info['annotated_dimensions']['bottom_center'].append((left_cut, right_cut))
            bars_info['info'].append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut, tie_info])))
        # Top right first order bar
        if sheet.read_cell('T101') != 0 or sheet.read_cell('T103') != 0:
            label = ""
            label_1 = str(int(sheet.read_cell('T101'))) + '%%C' + sheet.read_cell('V101')
            label_2 = str(int(sheet.read_cell('T103'))) + '%%C' + sheet.read_cell('V103')
            if sheet.read_cell('T101') != 0:
                label = label + label_1 + (' + ' if sheet.read_cell('T103') != 0 else '')
            if sheet.read_cell('T103') != 0:
                label = label + label_2
            case = 3
            side = 1
            order = 1
            left_cut = round(min(sheet.read_cell('F239') - sheet.read_cell('F257'),
                                 sheet.read_cell('F239') - sheet.read_cell('F257')), 2)
            right_cut = round(max(sheet.read_cell('F263') - sheet.read_cell('F257'),
                                  sheet.read_cell('F264') - sheet.read_cell('F257')), 2)
            tie_info = [[label_1, label_2], False, [False, False]]
            tie_info[1] = True if sheet.read_cell('AA99') else False
            tie_info[2][0] = True if sheet.read_cell('W101') else False
            tie_info[2][1] = True if sheet.read_cell('W103') else False
            bars_info['quantity']['top_right'] += 1
            bars_info['annotated_dimensions']['top_right'].append(left_cut)
            bars_info['info'].append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut, tie_info])))
        # Top right second order bar
        if sheet.read_cell('X103') != 0 or sheet.read_cell('X105') != 0:
            label = ""
            label_1 = str(int(sheet.read_cell('X103'))) + '%%C' + sheet.read_cell('Z103')
            label_2 = str(int(sheet.read_cell('X105'))) + '%%C' + sheet.read_cell('Z105')
            if sheet.read_cell('X103') != 0:
                label = label + label_1 + (' + ' if sheet.read_cell('X105') != 0 else '')
            if sheet.read_cell('X105') != 0:
                label = label + label_2
            case = 3
            side = 1
            order = 2
            left_cut = round(min(sheet.read_cell('F242') - sheet.read_cell('F257'),
                                 sheet.read_cell('F242') - sheet.read_cell('F257')), 2)
            right_cut = round(max(sheet.read_cell('F268') - sheet.read_cell('F257'),
                                  sheet.read_cell('F269') - sheet.read_cell('F257')), 2)
            tie_info = [[label_1, label_2], False, [False, False]]
            tie_info[1] = True if sheet.read_cell('AA101') else False
            tie_info[2][0] = True if sheet.read_cell('AA103') else False
            tie_info[2][1] = True if sheet.read_cell('AA105') else False
            bars_info['quantity']['top_right'] += 1
            bars_info['annotated_dimensions']['top_right'].append(left_cut)
            bars_info['info'].append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut, tie_info])))
        # Bottom right first order bar
        if sheet.read_cell('T118') != 0 or sheet.read_cell('T120') != 0:
            label = ""
            label_1 = str(int(sheet.read_cell('T118'))) + '%%C' + sheet.read_cell('V118')
            label_2 = str(int(sheet.read_cell('T120'))) + '%%C' + sheet.read_cell('V120')
            if sheet.read_cell('T118') != 0:
                label = label + label_1 + (' + ' if sheet.read_cell('T120') != 0 else '')
            if sheet.read_cell('T120') != 0:
                label = label + label_2
            case = 3
            side = -1
            order = 1
            left_cut = round(min(sheet.read_cell('F245') - sheet.read_cell('F257'),
                                 sheet.read_cell('F245') - sheet.read_cell('F257')), 2)
            right_cut = round(max(sheet.read_cell('F278') - sheet.read_cell('F257'),
                                  sheet.read_cell('F279') - sheet.read_cell('F257')), 2)
            tie_info = [[label_1, label_2], False, [False, False]]
            tie_info[1] = True if sheet.read_cell('AA123') else False
            tie_info[2][0] = True if sheet.read_cell('W118') else False
            tie_info[2][1] = True if sheet.read_cell('W120') else False
            bars_info['quantity']['bottom_right'] += 1
            bars_info['annotated_dimensions']['bottom_right'].append(left_cut)
            bars_info['info'].append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut, tie_info])))
        # Bottom right second order bar
        if sheet.read_cell('X116') != 0 or sheet.read_cell('X118') != 0:
            label = ""
            label_1 = str(int(sheet.read_cell('X116'))) + '%%C' + sheet.read_cell('Z116')
            label_2 = str(int(sheet.read_cell('X118'))) + '%%C' + sheet.read_cell('Z118')
            if sheet.read_cell('X116') != 0:
                label = label + label_1 + (' + ' if sheet.read_cell('X118') != 0 else '')
            if sheet.read_cell('X118') != 0:
                label = label + label_2
            case = 3
            side = -1
            order = 2
            left_cut = round(min(sheet.read_cell('F248') - sheet.read_cell('F257'),
                                 sheet.read_cell('F248') - sheet.read_cell('F257')), 2)
            right_cut = round(max(sheet.read_cell('F283') - sheet.read_cell('F257'),
                                  sheet.read_cell('F284') - sheet.read_cell('F257')), 2)
            tie_info = [[label_1, label_2], False, [False, False]]
            tie_info[1] = True if sheet.read_cell('AA121') else False
            tie_info[2][0] = True if sheet.read_cell('AA116') else False
            tie_info[2][1] = True if sheet.read_cell('AA118') else False
            bars_info['quantity']['bottom_right'] += 1
            bars_info['annotated_dimensions']['bottom_right'].append(left_cut)
            bars_info['info'].append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut, tie_info])))
        stirrups_info = {'differentiate': sheet.read_cell('I414'), 'diameters': {
            'l2r_diam': "%%C" + str(sheet.read_cell('L416')),
            'r2l_diam': "%%C" + str(sheet.read_cell('L424'))},
                         'quantity': {
                             'l2r_two_legged': sheet.read_cell('I416'),
                             'l2r_single_legged': sheet.read_cell('J416'),
                             'r2l_two_legged': sheet.read_cell('I424'),
                             'r2l_single_legged': sheet.read_cell('J424')},
                         'text': "",
                         'info': []}

        stirrups_text = str(sheet.read_cell('I416')) + ' (est. rect.) '
        stirrups_text = stirrups_text + (
            '+ ' + str(sheet.read_cell('J416')) + ' (gancho) ' if sheet.read_cell('J416') != 0 else '')
        stirrups_text = stirrups_text + '%%C' + sheet.read_cell('L416') + ': ' + sheet.read_cell('M431')
        if not stirrups_info['differentiate']:
            stirrups_text = stirrups_text + ' c/ext.'
        else:
            stirrups_text = stirrups_text + ' ----->    <----- '
            stirrups_text = stirrups_text + str(sheet.read_cell('I424')) + ' (est. rect.) '
            stirrups_text = stirrups_text + (
                '+ ' + str(sheet.read_cell('J424')) + ' (gancho) ' if sheet.read_cell('J424') != 0 else '')
            stirrups_text = stirrups_text + '%%C' + sheet.read_cell('L424') + ': ' + sheet.read_cell('U431')
        stirrups_info['text'] = stirrups_text

        stirrups_keys = ['side', 'quantity', 'spacing']
        l2r_row = 417
        while True:
            quantity = sheet.read_cell((l2r_row, 14))
            spacing = sheet.read_cell((l2r_row, 16))
            stirrups_info['info'].append(dict(zip(stirrups_keys, [0, quantity, spacing])))
            if sheet.read_cell((l2r_row, 13)) == 1:
                break
            l2r_row += 1
        r2l_row = 425
        while True:
            quantity = sheet.read_cell((r2l_row, 14))
            spacing = sheet.read_cell((r2l_row, 16))
            stirrups_info['info'].append(dict(zip(stirrups_keys, [1, quantity, spacing])))
            if sheet.read_cell((r2l_row, 13)) == 1:
                break
            r2l_row += 1
        return dict(zip(span_keys, [span_name, ls_info, rs_info, free_length, width, height, bars_info, stirrups_info]))


def aDouble(xyz):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, xyz)


def aDispatch(vObject):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, vObject)


class Point:
    def __init__(self, x: Union[int, float, np.ndarray, list, tuple, Vector], y: Union[int, float] = 0.0,
                 z: Union[int, float] = 0.0):
        if isinstance(x, (np.ndarray, list, tuple)):
            if len(x) == 3:
                self.x, self.y, self.z = x[0], x[1], x[2]
            elif len(x) == 2:
                self.x, self.y, self.z = x[0], x[1], z
            elif len(x) == 1:
                self.x, self.y, self.z = x[0], y, z
            else:
                raise Exception("Invalid number of coordinates")
        elif isinstance(x, (int, float)):
            self.x, self.y, self.z = x, y, z
        else:
            raise Exception("Integer or float expected")
        self.APoint = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (self.x, self.y, self.z))

    def distance2point(self, P0: Self) -> float:
        return math.sqrt((self.x - P0.x) ** 2 +
                         (self.y - P0.y) ** 2 +
                         (self.z - P0.z) ** 2)

    def distance2line(self, L0: 'Line') -> float:
        x, y, z = self.x, self.y, self.z
        A, B, C = L0.A, L0.B, L0.C
        return math.fabs(A * x + B * y + C) / math.sqrt(A ** 2 + B ** 2)

    def projection2line(self, L0: 'Line') -> Self:
        x, y, z = self.x, self.y, self.z
        m, b = L0.m, L0.b
        m2 = -1 / m
        b2 = y - m2 * x
        int_x = -(b - b2) / (m - m2)
        int_y = m2 * int_x + b2
        return Point(int_x, int_y, z)

    def rotation(self, c: Self, angle: Union[float, int]) -> Self:
        x, y, z = self.x, self.y, self.z
        cx, cy = c.x, c.y
        s = math.sin(angle)
        c = math.cos(angle)
        x = x - cx
        y = y - cy
        x_new = x * c - y * s
        y_new = x * s + y * c
        x = x_new + cx
        y = y_new + cy
        return Point(x, y, z)

    def interpolate2point(self, P0: Self, alpha: float) -> Self:
        x, y, z = self.x, self.y, self.z
        x0, y0, z0 = P0.x, P0.y, P0.z
        return Point(x0 * alpha + x * (1 - alpha), y0 * alpha + y * (1 - alpha), z0 * alpha + z * (1 - alpha))

    def is_collinear(self, L0: 'Line') -> bool:
        L1 = Line(L0.P0, self)
        if L0.is_same(L1):
            return True
        else:
            return False


class Line:
    def __init__(self, P0: Point, P1: Point):
        self.P0 = P0
        self.P1 = P1
        if self.P1.x == self.P0.x:
            self.m = None
            self.b = None
        else:
            self.m = (self.P1.y - self.P0.y) / (self.P1.x - self.P0.x)
            self.b = self.P0.y - self.m * self.P0.x
        self.A, self.B, self.C = self.line2general()

    def line2general(self) -> tuple[int | float, int | float, int | float]:
        if self.m is None:
            return 1, 0, -self.P0.x
        else:
            A, B, C = -self.m, 1, -self.b
        if A < 0:
            A, B, C = -A, -B, -C
        denA = Fraction(A).limit_denominator(1000).as_integer_ratio()[1]
        denC = Fraction(C).limit_denominator(1000).as_integer_ratio()[1]
        gcd = np.gcd(denA, denC)
        lcm = denA * denC / gcd
        A = A * lcm
        B = B * lcm
        C = C * lcm
        return A, B, C

    def intersect2line(self, L0: Self) -> Point:
        A, B, C = self.A, self.B, -self.C
        A0, B0, C0 = L0.A, L0.B, -L0.C
        if A * B0 - A0 * B == 0:
            raise Exception("Lines are parallel. There no intersection")
        else:
            return Point((C * B0 - C0 * B) / (A * B0 - A0 * B), (A * C0 - A0 * C) / (A * B0 - A0 * B))

    def mid_point(self) -> Point:
        return self.P0.interpolate2point(self.P1, 0.5)

    def is_parallel(self, L0: Self) -> bool:
        A, B, C = self.A, self.B, self.C
        A0, B0, C0 = L0.A, L0.B, L0.C
        if A0 != 0 and B0 != 0:
            if A / A0 == B / B0:
                return True
            else:
                return False
        else:
            if (A == 0 and A0 == 0) or (B == 0 and B0 == 0):
                return True
            else:
                return False

    def is_same(self, L0: Self) -> bool:
        A, B, C = self.A, self.B, self.C
        A0, B0, C0 = L0.A, L0.B, L0.C
        if A0 != 0 and B0 != 0 and C0 != 0:
            if A / A0 == B / B0 and B / B0 == C / C0:
                return True
            else:
                return False
        else:
            if (A == 1 and A0 == 1 and C == C0) or (B == 1 and B0 == 1 and C == C0):
                return True
            else:
                return False


class CAD:
    def __init__(self):
        self.acad = win32com.client.Dispatch("AutoCAD.Application")
        self.acad.Visible = True
        self.acad.Documents.Add()
        time.sleep(3)
        self.acadDoc = self.acad.ActiveDocument
        self.acadModel = self.acadDoc.ModelSpace
        self.objects_list = []
        self.selection_set = self.acadDoc.ActiveSelectionSet
        self.selected_objects = []
        self.layers = {}
        self.create_new_layer('LCM-TRAZO', 7)
        self.create_new_layer('LCM-ACERO', 4)
        self.create_new_layer('LCM-ESTRIBOS', 1)
        self.create_new_layer('LCM-TEXTOS', 3)
        self.create_new_layer('LCM-COTAS', 1)
        self.acadDoc.ActiveLayer = self.layers['LCM-TRAZO']
        self.create_new_dim_style('PRISMA 1-25')

    def create_new_dim_style(self, name: str = "1-100"):
        new_style = self.acad.ActiveDocument.DimStyles.Add(name)
        self.acadDoc.SetVariable("DIMDLE", 0.20)
        self.acadDoc.SetVariable("DIMDLI", 0.20)
        self.acadDoc.SetVariable("DIMEXE", 0.20)
        self.acadDoc.SetVariable("DIMEXO", 0.20)
        self.acadDoc.SetVariable("DIMBLK", 'ArchTick')
        self.acadDoc.SetVariable("DIMBLK1", 'ArchTick')
        self.acadDoc.SetVariable("DIMBLK2", 'ArchTick')
        self.acadDoc.SetVariable("DIMLDRBLK", 'ArchTick')
        self.acadDoc.SetVariable("DIMASZ", 0.25)
        self.acadDoc.SetVariable("DIMCEN", 0.09)
        self.acadDoc.SetVariable("DIMTXT", 0.25)
        self.acadDoc.SetVariable("DIMTAD", 2)
        self.acadDoc.SetVariable("DIMGAP", 0.1)
        self.acadDoc.SetVariable("DIMTMOVE", 2)
        self.acadDoc.SetVariable("DIMSCALE", 0.25)
        self.acadDoc.SetVariable("DIMDSEP", '.')
        self.acadDoc.SetVariable("DIMRND", 0.00)
        self.acadDoc.SetVariable("DIMZIN", 5)
        new_style.CopyFrom(self.acadDoc)
        self.acadDoc.ActiveDimStyle = new_style

    def create_new_layer(self, name: str, color_num: int = 1, line_type: str = 'Continuous',
                         line_weight: str = 'Default'):
        new_layer = self.acadDoc.Layers.Add(name)
        new_layer.color = color_num
        try:
            self.acadDoc.Linetypes.Load(line_type, 'acadiso.lin')
        except Exception:
            pass
        finally:
            new_layer.LineType = line_type
        if line_weight != 'Default':
            new_layer.LineWeight = line_weight
        self.layers[name] = new_layer

    def draw_beam(self, beam_info: dict, base_point: Vector = None):
        if not base_point:
            base_point = [0.0, 0.0]
        left_edge_width = beam_info['spans_info'][0]['left_support_info'][0]
        left_edge_type = beam_info['spans_info'][0]['left_support_info'][1]
        left_height = beam_info['spans_info'][0]['height']
        if left_edge_type == "Col/Pl":
            self.draw_line_by_points([base_point[0] - left_edge_width / 2, base_point[1] - left_height - 0.5],
                                     [base_point[0] - left_edge_width / 2, base_point[1] + 0.5])
        else:
            self.draw_line_by_points([base_point[0] - left_edge_width / 2, base_point[1] - left_height],
                                     [base_point[0] - left_edge_width / 2, base_point[1]])
        for span_info in beam_info['spans_info']:
            # w = span_info['width']  # width
            h = span_info['height']  # height
            fl = span_info['free_length']  # free length
            left_shw = span_info['left_support_info'][0] * 0.5  # half of left_support_width
            right_shw = span_info['right_support_info'][0] * 0.5  # half of right_support_width
            left_face = base_point[0] + left_shw
            right_face = base_point[0] + left_shw + fl
            self.draw_line_by_points([left_face, base_point[1]], [right_face, base_point[1]])
            left_edge_type = span_info['left_support_info'][1]
            if left_edge_type == "Viga":
                self.draw_line_by_points([left_face, base_point[1]],
                                         [left_face, base_point[1] - 0.5 * h])
            elif left_edge_type == "Col/Pl":
                self.draw_line_by_points([left_face, base_point[1]],
                                         [left_face, base_point[1] + 0.5])
            right_edge_type = span_info['right_support_info'][1]
            if right_edge_type == "Viga":
                self.draw_line_by_points([right_face, base_point[1]],
                                         [right_face, base_point[1] - 0.5 * h])
            elif right_edge_type == "Col/Pl":
                self.draw_line_by_points([right_face, base_point[1]],
                                         [right_face, base_point[1] + 0.5])
            self.select_last(3)
            self.mirror([left_face, base_point[1] - 0.5 * h], [right_face, base_point[1] - 0.5 * h])
            self.draw_linear_dimension([left_face, base_point[1] - h - 0.5],
                                       [right_face, base_point[1] - h - 0.5], -0.25)
            if left_shw != 0:
                if left_edge_type == "Col/Pl":
                    self.draw_concrete_extension([base_point[0] - left_shw,  base_point[1] + 0.5],
                                                 [base_point[0] + left_shw,  base_point[1] + 0.5])
                    self.select_last(5)
                    self.copy([0, 0.5], [0, -h - 0.5])
                else:
                    self.draw_line_by_points([base_point[0] - left_shw, base_point[1]],
                                             [base_point[0] + left_shw, base_point[1]])
                    self.select_last()
                    self.copy([0, 0], [0, -h])
                self.draw_linear_dimension([base_point[0] - left_shw, base_point[1] - h - 0.5],
                                           [base_point[0] + left_shw, base_point[1] - h - 0.5], -0.25)
            for bar_data in span_info['bars_info']['info']:
                tie_info = bar_data['tie_info']
                order = bar_data['order']
                side = bar_data['side']
                if bar_data['case'] == 0:
                    if tie_info[1]:
                        bar_data['left_cut'] = bar_data['left_cut'] / 2
                    if tie_info[3]:
                        bar_data['right_cut'] = bar_data['right_cut'] / 2
                    bar_data['left_annotation'] = None
                    bar_data['right_annotation'] = None
                    bar_data['text_position'] = 0
                elif bar_data['case'] == 1:
                    side = 'top_left' if side == 1 else 'bottom_left'
                    if tie_info[1]:
                        bar_data['left_cut'] = bar_data['left_cut'] / 2
                    if order == 1 and span_info['bars_info']['quantity'][side] == 2:
                        bar_data['left_annotation'] = span_info['bars_info']['annotated_dimensions'][side][1]
                        bar_data['right_annotation'] = bar_data['right_cut']
                    else:
                        bar_data['left_annotation'] = 0.0
                        bar_data['right_annotation'] = bar_data['right_cut']
                    bar_data['text_position'] = 1
                elif bar_data['case'] == 2:
                    side = 'bottom_center'
                    if order == 2 and span_info['bars_info']['quantity'][side] == 2:
                        bar_data['left_annotation'] = span_info['bars_info']['annotated_dimensions'][side][0][0]
                        bar_data['right_annotation'] = span_info['bars_info']['annotated_dimensions'][side][0][1]
                        bar_data['text_position'] = 0
                    else:
                        bar_data['left_annotation'] = 0.0
                        bar_data['right_annotation'] = 0.0
                        if span_info['bars_info']['quantity'][side] == 2:
                            bar_data['text_position'] = -1
                        else:
                            bar_data['text_position'] = 0
                elif bar_data['case'] == 3:
                    side = 'top_right' if side == 1 else 'bottom_right'
                    if tie_info[1]:
                        bar_data['right_cut'] = bar_data['right_cut'] / 2
                    if order == 1 and span_info['bars_info']['quantity'][side] == 2:
                        bar_data['left_annotation'] = bar_data['left_cut']
                        bar_data['right_annotation'] = span_info['bars_info']['annotated_dimensions'][side][1]
                    else:
                        bar_data['left_annotation'] = bar_data['left_cut']
                        bar_data['right_annotation'] = 0.0
                    bar_data['text_position'] = -1
                self.draw_beam_longitudinal_bar(base_point[1] - h / 2, h, left_face, right_face, bar_data, )
            self.draw_text(span_info['span_name'],
                           Point((left_face + right_face) / 2, base_point[1] + 0.75), 0.10)
            self.draw_text(span_info['stirrups_info']['text'],
                           Point((left_face + right_face) / 2, base_point[1] - h - 0.4))
            base_point[0] += left_shw + fl + right_shw
        right_edge_width = beam_info['spans_info'][-1]['right_support_info'][0]
        right_edge_type = beam_info['spans_info'][-1]['right_support_info'][1]
        right_height = beam_info['spans_info'][-1]['height']
        if right_edge_type == "Col/Pl":
            self.draw_line_by_points([base_point[0] + right_edge_width / 2, base_point[1] - right_height - 0.5],
                                     [base_point[0] + right_edge_width / 2, base_point[1] + 0.5])
            if right_edge_width != 0:
                self.draw_concrete_extension([base_point[0] - right_edge_width / 2, base_point[1] + 0.5],
                                             [base_point[0] + right_edge_width / 2, base_point[1] + 0.5])
                self.select_last(5)
                self.copy([0, 0.5], [0, -right_height - 0.5])
                self.draw_linear_dimension([base_point[0] - right_edge_width / 2, base_point[1] - right_height - 0.5],
                                           [base_point[0] + right_edge_width / 2, base_point[1] - right_height - 0.5],
                                           -0.25)
        else:
            self.draw_line_by_points([base_point[0] + right_edge_width / 2, base_point[1] - right_height],
                                     [base_point[0] + right_edge_width / 2, base_point[1]])
            if right_edge_type == "Viga" and right_edge_width != 0:
                self.draw_line_by_points([base_point[0] - right_edge_width / 2, base_point[1]],
                                         [base_point[0] + right_edge_width / 2, base_point[1]])
                self.select_last()
                self.copy([0, 0], [0, -right_height])
                self.draw_linear_dimension([base_point[0] - right_edge_width / 2, base_point[1] - right_height - 0.5],
                                           [base_point[0] + right_edge_width / 2, base_point[1] - right_height - 0.5],
                                           -0.25)

    def draw_line(self, L0: Line, layer: str = 'LCM-TRAZO'):
        L1 = self.acadModel.AddLine(L0.P0.APoint, L0.P1.APoint)
        L1.layer = layer
        self.objects_list.append(L1)

    def draw_line_by_points(self, P0: Union[Point, list], P1: Union[Point, list], layer: str = 'LCM-TRAZO'):
        if not isinstance(P0, Point):
            P0 = Point(P0)
        if not isinstance(P1, Point):
            P1 = Point(P1)
        L1 = self.acadModel.AddLine(P0.APoint, P1.APoint)
        L1.layer = layer
        self.objects_list.append(L1)

    def draw_text(self, text: str, P0: Union[Point, list], TSize: float = 0.05, layer: str = 'LCM-TEXTOS',
                  alignment: int = 10, MText: bool = False, BoxWidth: float = 0):
        if not isinstance(P0, Point):
            P0 = Point(P0)
        if MText:
            T1 = self.acadModel.AddMText(P0.APoint, BoxWidth, text)
        else:
            T1 = self.acadModel.AddText(text, P0.APoint, TSize)
        T1.Layer = layer
        T1.HorizontalAlignment = 1
        T1.TextAlignmentPoint = P0.APoint
        T1.Alignment = alignment
        self.objects_list.append(T1)

    def draw_linear_dimension(self, P0: Union[Point, list], P1: Union[Point, list], text_offset: float = 0.25,
                              layer: str = 'LCM-COTAS'):
        if not isinstance(P0, Point):
            P0 = Point(P0)
        if not isinstance(P1, Point):
            P1 = Point(P1)
        P2 = Point((P0.x + P1.x) / 2 + (text_offset if P0.x == P1.x else 0),
                   (P0.y + P1.y) / 2 + (text_offset if P0.y == P1.y else 0))
        D1 = self.acadModel.AddDimRotated(P0.APoint, P1.APoint, P2.APoint, 0 if P0.y == P1.y else math.pi / 2)
        D1.Layer = layer
        self.objects_list.append(D1)

    def draw_concrete_extension(self, P0: Union[Point, list], P1: Union[Point, list], fixed_height=0.2, ratio=0.0):
        if not isinstance(P0, Point):
            P0 = Point(P0)
        if not isinstance(P1, Point):
            P1 = Point(P1)
        h = fixed_height
        d = P1.distance2point(P0)
        if P0.x == P1.x:
            angle = 0.5 * math.pi
        else:
            angle = math.atan((P1.y - P0.y) / (P1.x - P0.x))
        if ratio != 0:
            h = ratio * d
        P0p = Point(P0.x + 0.5 * (d - 0.5 * h) * math.cos(angle), P0.y + 0.5 * (d - 0.5 * h) * math.sin(angle))
        P1p = Point(P1.x - 0.5 * (d - 0.5 * h) * math.cos(angle), P1.y - 0.5 * (d - 0.5 * h) * math.sin(angle))
        P2t = Point(P0.x + 0.5 * d * math.cos(angle) + 0.5 * h * math.cos(angle + 0.5 * math.pi),
                    P0.y + 0.5 * d * math.sin(angle) + 0.5 * h * math.sin(angle + 0.5 * math.pi))
        P2b = Point(P1.x - 0.5 * d * math.cos(angle) - 0.5 * h * math.cos(angle + 0.5 * math.pi),
                    P1.y - 0.5 * d * math.sin(angle) - 0.5 * h * math.sin(angle + 0.5 * math.pi))
        self.draw_line_by_points(P0, P0p)
        self.draw_line_by_points(P0p, P2t)
        self.draw_line_by_points(P2t, P2b)
        self.draw_line_by_points(P2b, P1p)
        self.draw_line_by_points(P1p, P1)

    def draw_beam_longitudinal_bar(self, beam_middle: float, beam_height: float,
                                   left_face: float, right_face: float,
                                   bar_data: dict):
        label = bar_data['label']
        case = bar_data['case']
        side = bar_data['side']
        order = bar_data['order']
        lc = bar_data['left_cut']
        rc = bar_data['right_cut']
        left_annotation = bar_data['left_annotation']
        right_annotation = bar_data['right_annotation']
        text_position = bar_data['text_position']
        tie_info = bar_data['tie_info']
        edge_offset = 0.05 + 0.05 * order
        bhh = beam_height / 2
        if case == 0:
            self.draw_line_by_points(Point(left_face + lc, beam_middle + (bhh - edge_offset) * side),
                                     Point(right_face + rc, beam_middle + (bhh - edge_offset) * side),
                                     'LCM-ACERO')
            self.draw_text(label, Point((left_face + right_face) / 2, beam_middle + (bhh + edge_offset) * side))
            if any(tie_info[2]):
                self.draw_line_by_points(Point(left_face + lc, beam_middle + (bhh - edge_offset) * side),
                                         Point(left_face + lc, beam_middle + (bhh - edge_offset) * side - 0.25 * side),
                                         'LCM-ACERO')
            if any(tie_info[4]):
                self.draw_line_by_points(Point(right_face + rc, beam_middle + (bhh - edge_offset) * side),
                                         Point(right_face + rc, beam_middle + (bhh - edge_offset) * side - 0.25 * side),
                                         'LCM-ACERO')
        elif case == 1:
            self.draw_line_by_points(Point(left_face + lc, beam_middle + (bhh - edge_offset) * side),
                                     Point(left_face + rc, beam_middle + (bhh - edge_offset) * side),
                                     'LCM-ACERO')
            self.draw_text(label, Point(left_face + rc, beam_middle + (bhh - edge_offset - 0.05) * side),
                           alignment=11)
            self.draw_linear_dimension(Point(left_face + left_annotation, beam_middle + bhh * side),
                                       Point(left_face + right_annotation, beam_middle + bhh * side),
                                       text_offset=0.25 * side)
            if any(tie_info[2]):
                self.draw_line_by_points(Point(left_face + lc, beam_middle + (bhh - edge_offset) * side),
                                         Point(left_face + lc, beam_middle + (bhh - edge_offset) * side - 0.25 * side),
                                         'LCM-ACERO')
        elif case == 2:
            self.draw_line_by_points(Point(left_face + lc, beam_middle + (bhh - edge_offset) * side),
                                     Point(right_face + rc, beam_middle + (bhh - edge_offset) * side),
                                     'LCM-ACERO')
            if text_position == -1:
                self.draw_text(label, Point(left_face + lc, beam_middle + (bhh - edge_offset - 0.05) * side),
                               alignment=9)
            else:
                self.draw_text(label, Point((left_face + right_face) / 2,
                                            beam_middle + (bhh - edge_offset - 0.05) * side))
            self.draw_linear_dimension(Point(left_face + left_annotation, beam_middle + bhh * side),
                                       Point(left_face + lc, beam_middle + bhh * side),
                                       text_offset=0.25 * side)
            self.draw_linear_dimension(Point(right_face + right_annotation, beam_middle + bhh * side),
                                       Point(right_face + rc, beam_middle + bhh * side),
                                       text_offset=0.25 * side)
        elif case == 3:
            self.draw_line_by_points(Point(right_face + lc, beam_middle + (bhh - edge_offset) * side),
                                     Point(right_face + rc, beam_middle + (bhh - edge_offset) * side),
                                     'LCM-ACERO')
            self.draw_text(label, Point(right_face + lc, beam_middle + (bhh - edge_offset - 0.05) * side),
                           alignment=9)
            self.draw_linear_dimension(Point(right_face + left_annotation, beam_middle + bhh * side),
                                       Point(right_face + right_annotation, beam_middle + bhh * side),
                                       text_offset=0.25 * side)
            if any(tie_info[2]):
                self.draw_line_by_points(Point(right_face + rc, beam_middle + (bhh - edge_offset) * side),
                                         Point(right_face + rc, beam_middle + (bhh - edge_offset) * side - 0.25 * side),
                                         'LCM-ACERO')

    def select_last(self, num_objects=1, selection_offset=0):
        selection = []
        for i in range(num_objects):
            obj = self.objects_list[-1 - selection_offset - i]
            selection.append(obj)
        self.selected_objects = selection

    def select_all(self):
        self.deselect_all()
        self.selection_set.Select(5)
        for i in range(self.selection_set.Count):
            self.selected_objects.append(self.selection_set.Item(i))

    def deselect_all(self):
        self.selected_objects = []
        self.selection_set.Clear()

    def erase_all(self):
        self.select_all()
        self.selection_set.Erase()

    def move(self, P0, P1):
        P0 = Point(P0).APoint
        P1 = Point(P1).APoint
        for obj in self.selected_objects:
            obj.Move(P0, P1)

    def move_all(self, P0, P1):
        P0 = Point(P0).APoint
        P1 = Point(P1).APoint
        # self.select_all()
        for obj in self.acadModel:
            obj.Move(P0, P1)

    def copy(self, P0, P1):
        P0 = Point(P0).APoint
        P1 = Point(P1).APoint
        for obj in self.selected_objects:
            copy = obj.Copy()
            self.objects_list.append(copy)
            copy.Move(P0, P1)

    def mirror(self, P0, P1):
        P0 = Point(P0).APoint
        P1 = Point(P1).APoint
        for obj in self.selected_objects:
            mirror = obj.Mirror(P0, P1)
            self.objects_list.append(mirror)

    def array(self, rows_number, columns_number, rows_spacing, columns_spacing, levels_num=1, levels_sp=0):
        for obj in self.selected_objects:
            try:
                obj.ArrayRectangular(rows_number, columns_number, levels_num, rows_spacing, columns_spacing, levels_sp)
            except KeyError:
                pass
            finally:
                pass
                # self.list_new_objects(rows_number * columns_number * levels_num - 1)

    def zoom_all(self):
        self.acad.ZoomExtents()

    def list_new_objects(self, num_objects):
        count = 0
        for obj in self.acadModel:
            self.objects_list.append(obj)
            count += 1
            if count == num_objects:
                break


if __name__ == '__main__':
    assistant = Assistant()
    list_indexes = assistant.get_sheets2draw()
    list_indexes = np.array(list_indexes) + 6
    assistant.download_excel_beams_info(list_indexes)
    assistant.create_json_file()
    assistant.read_json_file()
    draftsman = CAD()
    b_point = [0, 0]
    for name, info in assistant.jsonDict.items():
        draftsman.draw_beam(info, base_point=b_point)
        b_point[0] = 0
        b_point[1] += 5
    draftsman.zoom_all()
