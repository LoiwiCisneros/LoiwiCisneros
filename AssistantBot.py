import json
import openpyxl
from tkinter import filedialog


class Assistant:
    def __init__(self, jsonFileName='Beams_info', xlsxFilePath=None):
        self.fileName = jsonFileName
        if xlsxFilePath is None:
            xlsxFilePath = filedialog.askopenfilename(filetypes=(("Excel files", "*xlsx"), ("Excel files", "*xlsm")))
        self.wb = openpyxl.load_workbook(xlsxFilePath, read_only=True, data_only=True, keep_vba=False)

    def get_variable_value(self, variable):
        with open(self.fileName) as jsonFile:
            dictionary = json.load(jsonFile)
            value = dictionary[variable]
        return value

    def set_variable_value(self, variable, value):
        with open(self.fileName) as jsonFile:
            dictionary = json.load(jsonFile)
        with open(self.fileName, 'w') as jsonFile:
            dictionary[variable] = value
            json.dump(dictionary, jsonFile)

    def set_default_variable(self, key, value):
        with open(self.fileName) as jsonFile:
            dictionary = json.load(jsonFile)
        dictionary.setdefault(str(key), value)
        with open(self.fileName, 'w') as jsonFile:
            json.dump(dictionary, jsonFile)

    def reset_values(self):
        default_values = {
            "Beams_info": {}
        }
        dictionary = default_values.get(self.fileName)
        with open(self.fileName, 'w') as jsonFile:
            json.dump(dictionary, jsonFile)

    def download_excel_beams_info(self, star_index: int = 3, last_index: int = None):
        beams_nums = []
        self.reset_values()
        for index in range(star_index, (len(self.wb.sheetnames) if last_index is None else last_index)):
            span_info = self.download_excel_span_info(index)
            beam_name = span_info['span_name'].split('-')[0] + span_info['span_name'].split('-')[1]
            if span_info['span_name'].split('-')[1] not in beams_nums:
                beams_nums.append(span_info['span_name'].split('-')[1])
                self.set_default_variable(beam_name, {
                    "beam_name": beam_name,
                    "spans_num": 0,
                    "spans_info": []
                })
            beam_dict = self.get_variable_value(beam_name)
            spans_num = beam_dict['spans_num']
            spans_info = beam_dict['spans_info']
            spans_info.append(span_info)
            beam_dict.update({
                "beam_name": beam_name,
                "spans_num": spans_num + 1,
                "spans_info": spans_info
            })
            self.set_variable_value(beam_name, beam_dict)

    def download_excel_span_info(self, index):
        ws = self.wb[self.wb.sheetnames[index]]
        span_keys = ['span_name', 'left_support_width', 'free_length', 'right_support_width', 'width', 'height',
                     'bars_info', 'stirrups_info']
        beam_name = ws.title
        left_support_width = ws['M78'].value
        free_length = ws['M79'].value
        right_support_width = ws['M80'].value
        width = ws['Q78'].value / 100
        height = ws['Q79'].value / 100

        bars_info = []
        bars_keys = ['label', 'case', 'side', 'order', 'left_cut', 'right_cut']
        # Top long bar
        if ws['O90'].value != 0 or ws['O91'].value != 0:
            label = ''
            if ws['O90'].value != 0:
                label = label + str(ws['O90'].value) + '%%C' + ws['Q90'].value + (' + ' if ws['O91'].value != 0 else '')
            if ws['O91'].value != 0:
                label = label + str(ws['O91'].value) + '%%C' + ws['Q91'].value
            case = 0
            side = 1
            order = 0
            left_cut = 0
            right_cut = 0
            bars_info.append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut])))
        # Bottom long bar
        if ws['O135'].value != 0 or ws['O136'].value != 0:
            label = ''
            if ws['O135'].value != 0:
                label = label + str(ws['O135'].value) + '%%C' + ws['Q135'].value + \
                        (' + ' if ws['O136'].value != 0 else '')
            if ws['O136'].value != 0:
                label = label + str(ws['O136'].value) + '%%C' + ws['Q136'].value
            case = 0
            side = -1
            order = 0
            left_cut = 0
            right_cut = 0
            bars_info.append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut])))
        # Top left first order bar
        if ws['J101'].value != 0 or ws['J103'].value != 0:
            label = ''
            if ws['J101'].value != 0:
                label = label + str(ws['J101'].value) + '%%C' + ws['L101'].value + \
                        (' + ' if ws['J103'].value != 0 else '')
            if ws['J103'].value != 0:
                label = label + str(ws['J103'].value) + '%%C' + ws['L103'].value
            case = 1
            side = 1
            order = 1
            left_cut = 0
            right_cut = ws['H92'].value
            bars_info.append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut])))
        # Top left second order bar
        if ws['F103'].value != 0 or ws['F105'].value != 0:
            label = ''
            if ws['F103'].value != 0:
                label = label + str(ws['F103'].value) + '%%C' + ws['H103'].value + \
                        (' + ' if ws['F105'].value != 0 else '')
            if ws['F105'].value != 0:
                label = label + str(ws['F105'].value) + '%%C' + ws['H105'].value
            case = 1
            side = 1
            order = 2
            left_cut = 0
            right_cut = ws['H92'].value - 0.4
            bars_info.append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut])))
        # Bottom left first order bar
        if ws['J118'].value != 0 or ws['J120'].value != 0:
            label = ''
            if ws['J118'].value != 0:
                label = label + str(ws['J118'].value) + '%%C' + ws['L118'].value + \
                        (' + ' if ws['J120'].value != 0 else '')
            if ws['J120'].value != 0:
                label = label + str(ws['J120'].value) + '%%C' + ws['L120'].value
            case = 1
            side = -1
            order = 1
            left_cut = 0
            right_cut = ws['H128'].value
            bars_info.append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut])))
        # Bottom left second order bar
        if ws['F116'].value != 0 or ws['F118'].value != 0:
            label = ''
            if ws['F116'].value != 0:
                label = label + str(ws['F116'].value) + '%%C' + ws['H116'].value + \
                        (' + ' if ws['F118'].value != 0 else '')
            if ws['F118'].value != 0:
                label = label + str(ws['F118'].value) + '%%C' + ws['H118'].value
            case = 1
            side = -1
            order = 2
            left_cut = 0
            right_cut = ws['H128'].value - 0.4
            bars_info.append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut])))
        # Bottom central first order bar
        if ws['O116'].value != 0 or ws['O118'].value != 0:
            label = ''
            if ws['O116'].value != 0:
                label = label + str(ws['O116'].value) + '%%C' + ws['Q116'].value + \
                        (' + ' if ws['O118'].value != 0 else '')
            if ws['O118'].value != 0:
                label = label + str(ws['O118'].value) + '%%C' + ws['Q118'].value
            case = 2
            side = -1
            order = 1
            left_cut = ws['J132'].value
            right_cut = ws['V132'].value
            bars_info.append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut])))
        # Bottom central second order bar
        if ws['O108'].value != 0 or ws['O110'].value != 0:
            label = ''
            if ws['O108'].value != 0:
                label = label + str(ws['O108'].value) + '%%C' + ws['Q108'].value + \
                        (' + ' if ws['O110'].value != 0 else '')
            if ws['O110'].value != 0:
                label = label + str(ws['O110'].value) + '%%C' + ws['Q110'].value
            case = 2
            side = -1
            order = 2
            left_cut = ws['J132'].value + 0.4
            right_cut = ws['V132'].value + 0.4
            bars_info.append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut])))
        # Top right first order bar
        if ws['T101'].value != 0 or ws['T103'].value != 0:
            label = ''
            if ws['T101'].value != 0:
                label = label + str(ws['T101'].value) + '%%C' + ws['V101'].value + \
                        (' + ' if ws['T103'].value != 0 else '')
            if ws['T103'].value != 0:
                label = label + str(ws['T103'].value) + '%%C' + ws['V103'].value
            case = 3
            side = 1
            order = 1
            left_cut = ws['X92'].value
            right_cut = 0
            bars_info.append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut])))
        # Top right second order bar
        if ws['X103'].value != 0 or ws['X105'].value != 0:
            label = ''
            if ws['X103'].value != 0:
                label = label + str(ws['X103'].value) + '%%C' + ws['Z103'].value + \
                        (' + ' if ws['X105'].value != 0 else '')
            if ws['X105'].value != 0:
                label = label + str(ws['X105'].value) + '%%C' + ws['Z105'].value
            case = 3
            side = 1
            order = 2
            left_cut = ws['X92'].value - 0.4
            right_cut = 0
            bars_info.append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut])))
        # Bottom right first order bar
        if ws['T118'].value != 0 or ws['T120'].value != 0:
            label = ''
            if ws['T118'].value != 0:
                label = label + str(ws['T118'].value) + '%%C' + ws['V118'].value + \
                        (' + ' if ws['T120'].value != 0 else '')
            if ws['T120'].value != 0:
                label = label + str(ws['T120'].value) + '%%C' + ws['V120'].value
            case = 3
            side = -1
            order = 1
            left_cut = ws['X128'].value
            right_cut = 0
            bars_info.append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut])))
        # Bottom right second order bar
        if ws['X116'].value != 0 or ws['X118'].value != 0:
            label = ''
            if ws['X116'].value != 0:
                label = label + str(ws['X116'].value) + '%%C' + ws['Z116'].value + \
                        (' + ' if ws['X118'].value != 0 else '')
            if ws['X118'].value != 0:
                label = label + str(ws['X118'].value) + '%%C' + ws['Z118'].value
            case = 3
            side = -1
            order = 2
            left_cut = ws['X128'].value - 0.4
            right_cut = 0
            bars_info.append(dict(zip(bars_keys, [label, case, side, order, left_cut, right_cut])))

        stirrups_info = str(ws['I400'].value) + ' (est. rect.) '
        stirrups_info = stirrups_info + ('+ ' + str(ws['J400'].value) + ' (gancho) ' if ws['J400'].value != 0 else '')
        stirrups_info = stirrups_info + '%%C' + ws['L400'].value + ': ' + ws['N416'].value
        dif_stirrups = ws['Q416'].value == 'c/ext.'
        if dif_stirrups:
            stirrups_info = stirrups_info + ' c/ext.'
        else:
            stirrups_info = stirrups_info + ' ----->    <----- '
            stirrups_info = stirrups_info + str(ws['I409'].value) + ' (est. rect.) '
            stirrups_info = stirrups_info + (
                '+ ' + str(ws['J409'].value) + ' (gancho) ' if ws['J409'].value != 0 else '')
            stirrups_info = stirrups_info + '%%C' + ws['L409'].value + ': ' + ws['V416'].value
        return dict(zip(span_keys, [beam_name, left_support_width, free_length, right_support_width, width, height,
                                    bars_info, stirrups_info]))


if __name__ == '__main__':
    ast = Assistant()
