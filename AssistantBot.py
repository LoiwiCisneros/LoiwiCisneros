import json
import os.path
import re
import tkinter as tk
from tkinter import ttk, filedialog
import numpy as np
import pandas as pd
from typing import Union, TypeAlias, Any

RowColumnNum: TypeAlias = tuple[int, int]


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

    def download_excel_beams_info(self, list_indexes: Union[list, tuple, np.ndarray]):
        for index in list_indexes:
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


if __name__ == '__main__':
    ast = Assistant()
