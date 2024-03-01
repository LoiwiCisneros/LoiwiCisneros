import win32com.client
import pythoncom
# import comtypes.client
import math
from fractions import Fraction
import numpy as np
import time
from AssistantBot import Assistant
from operator import itemgetter
from typing import Union, TypeAlias
from typing import Annotated
from typing import Self
from typing import Any
import json

Vector: TypeAlias = list[Union[int, float]]

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


# def get_dev_length(D, tie_case):
#     lengths = AssistantBot.get_variable_value("DEV_LENGTHS")
#     if tie_case != 2:
#         # if case == 0:
#         #     ld = 0
#         # else:
#         ld = lengths.get(str(tie_case) + str(D))
#         return ld
#     else:
#         [ld, tie_long] = lengths.get(str(tie_case) + str(D))
#         return [ld, tie_long]


# def determine_tie_case(bar_case, bar_restriction, side):
#     if side == 1:
#         case = 0
#         momentum_sign = get_envelope_momentum_signs(bar_case)[0]
#     elif side == -1:
#         case = 1
#         momentum_sign = get_envelope_momentum_signs(bar_case)[1]
#
#     return case
#
#
# def get_envelope_momentum_signs(bar_case):
#     AssistantBot.get_variable_value("ACTUAL_BEAM_FORCES_INFO").get(str(bar_case))
#     return sign


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
            right_edge_type = span_info['left_support_info'][1]
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
                self.draw_linear_dimension([base_point - left_shw, -h - 0.5],
                                           [base_point + left_shw, -h - 0.5], -0.25)
            for bar_data in span_info['bars_info']['info']:
                self.draw_beam_longitudinal_bar(h / 2, left_face, right_face, bar_data)
            self.draw_text(span_info['span_name'], Point((left_face + right_face) / 2, 0.75), 0.10)
            self.draw_text(span_info['stirrups_info']['text'], Point((left_face + right_face) / 2, -h - 0.4))
            base_point += left_shw + fl + right_shw
        right_edge_width = beam_info['spans_info'][-1]['right_support_info'][0]
        right_edge_type = beam_info['spans_info'][-1]['right_support_info'][1]
        right_height = beam_info['spans_info'][-1]['height']
        if right_edge_width != 0:
            if right_edge_type == "Col/Pl":
                self.draw_line_by_points([base_point[0] + right_edge_width / 2, base_point[1] - right_height - 0.5],
                                         [base_point[0] + right_edge_width / 2, base_point[1] + 0.5])
                self.draw_concrete_extension([base_point[0] - right_edge_width / 2, 0.5],
                                             [base_point[0] + right_edge_width / 2, 0.5])
                self.select_last(5)
                self.copy([0, 0.5], [0, -right_height - 0.5])
            else:
                self.draw_line_by_points([base_point[0] + right_edge_width / 2, base_point[1] - right_height],
                                         [base_point[0] + right_edge_width / 2, base_point[1]])
                self.select_last()
                self.copy([0, 0], [0, -right_height])
            self.draw_linear_dimension([base_point[0] - right_edge_width / 2, base_point[1] - right_height - 0.5],
                                       [base_point[0] + right_edge_width / 2, base_point[1] - right_height - 0.5],
                                       -0.25)

    def draw_column(self):
        pass

    def draw_wall(self):
        pass

    def draw_footing(self):
        pass

    def draw_point(self, P0, layer='A-TRAZO'):
        if not isinstance(P0, Point):
            P0 = Point(P0)
        P0 = self.acadModel.AddPoint(P0.APoint)
        P0.layer = layer
        self.objects_list.append(P0)

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

    def draw_polyline(self, points, layer='LCM-TRAZO'):
        points = aDouble(points)
        PL1 = self.acadModel.AddPolyline(points)
        PL1.layer = layer
        self.objects_list.append(PL1)

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

    def draw_beam_longitudinal_bar(self, beam_middle: float, left_face: float, right_face: float, bar_data: dict):
        label = bar_data['label']
        case = bar_data['case']
        side = bar_data['side']
        order = bar_data['order']
        left_cut = bar_data['left_cut']
        lc = min(left_cut)
        right_cut = bar_data['right_cut']
        rc = max(right_cut)
        tie_info = bar_data['tie_info']
        edge_offset = 0.05 + 0.05 * order
        if case == 0:
            self.draw_line_by_points(Point(left_face + lc, -beam_middle + (beam_middle - edge_offset) * side),
                                     Point(right_face + rc, -beam_middle + (beam_middle - edge_offset) * side),
                                     'LCM-ACERO')
            # if any(tie_info[2]):
            #     db1 = tie_info[0][0].split('C')[1].split('/')[0]
            #     # db = max(tie_info[0][0] * (0 if ),)
            #     self.draw_tie_long_bar(Point(left_cut[0 if tie_info[2][0] else 1],
            #                                  -beam_middle + (beam_middle - edge_offset) * side),
            #                            db)
            # if any(tie_info[4]):
            #     self.draw_tie_long_bar()
            self.draw_text(label, Point((left_face + right_face) / 2,
                                        -beam_middle + (beam_middle + edge_offset) * side))
        elif case == 1:
            self.draw_line_by_points(Point(left_face + lc, -beam_middle + (beam_middle - edge_offset) * side),
                                     Point(left_face + rc, -beam_middle + (beam_middle - edge_offset) * side),
                                     'LCM-ACERO')
            self.draw_text(label, Point(left_face + rc - 0.1,
                                        -beam_middle + (beam_middle - edge_offset - 0.05) * side))
            self.draw_linear_dimension(Point(left_face, -beam_middle + beam_middle * side),
                                       Point(left_face + rc, -beam_middle + beam_middle * side),
                                       text_offset=0.25 * side)
        elif case == 2:
            self.draw_line_by_points(Point(left_face + lc, -beam_middle + (beam_middle - edge_offset) * side),
                                     Point(right_face + rc, -beam_middle + (beam_middle - edge_offset) * side),
                                     'LCM-ACERO')
            self.draw_text(label, Point(left_face + lc + 0.1,
                                        -beam_middle + (beam_middle - edge_offset - 0.05) * side))
            self.draw_linear_dimension(Point(left_face, -beam_middle + beam_middle * side),
                                       Point(left_face + lc, -beam_middle + beam_middle * side),
                                       text_offset=0.25 * side)
            self.draw_linear_dimension(Point(right_face, -beam_middle + beam_middle * side),
                                       Point(right_face + rc, -beam_middle + beam_middle * side),
                                       text_offset=0.25 * side)
        elif case == 3:
            self.draw_line_by_points(Point(right_face + lc, -beam_middle + (beam_middle - edge_offset) * side),
                                     Point(right_face + rc, -beam_middle + (beam_middle - edge_offset) * side),
                                     'LCM-ACERO')
            self.draw_text(label, Point(right_face + lc + 0.1,
                                        -beam_middle + (beam_middle - edge_offset - 0.05) * side))
            self.draw_linear_dimension(Point(right_face, -beam_middle + beam_middle * side),
                                       Point(right_face + lc, -beam_middle + beam_middle * side),
                                       text_offset=0.25 * side)
        # if left_con == 0:
        #     tie_case = determine_tie_case(bar_case, bar_restrictions[0], side)
        #     ld = get_dev_length(D, tie_case)
        #     ld = min(ld, bar_restrictions[0]-0.06-0.04*order)
        #     self.draw_tie_long_bar([ip[0]+left_cut, -ip[1]+(ip[1]-edge_offset)*side, 0], ld, tie_case, side, -1)
        # if right_con == 0:
        #     tie_case = determine_tie_case(bar_case, bar_restrictions[1], side)
        #     ld = get_dev_length(D, tie_case)
        #     ld = min(ld, bar_restrictions[1]-0.06-0.04*order)
        #     self.draw_tie_long_bar([ip[0] + right_cut, -ip[1] + (ip[1] - edge_offset) * side, 0], ld, tie_case, side,
        #                            1)

    def draw_tie_long_bar(self, P0: Point, db: list, tie: list, side=1):
        if db == "8mm":
            db = 0.08
        else:
            db = int(db.split('/')[0]) * 2.54 / 8
        P1 = Point(P0.x, P0.y - side * 16 * db)
        self.draw_line_by_points(P0, P1, 'A-ACERO')

    def select_last(self, num_objects=1, selection_offset=0):
        selection = []
        for i in range(num_objects):
            obj = self.objects_list[-1 - selection_offset - i]
            selection.append(obj)
        # self.select_all()
        # for i in range(num_objects):
        #     obj = self.selection_set.Item(i + selection_offset)
        #     selection.append(obj)
        # self.deselect_all()
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
        # self.select_last(num_objects)
        # for obj in self.selected_objects:
        #     self.objects_list.append(obj)


def is_point_in_triangle(vp, va, vb, vc, counter=1):
    vab = np.subtract(vb, va)
    vbc = np.subtract(vc, vb)
    vca = np.subtract(va, vc)
    vap = np.subtract(vp, va)
    vbp = np.subtract(vp, vb)
    vcp = np.subtract(vp, vc)

    cross1 = np.cross(vab, vap) * counter
    cross2 = np.cross(vbc, vbp) * counter
    cross3 = np.cross(vca, vcp) * counter

    if cross1 > 0 or cross2 > 0 or cross3 > 0:
        return False
    return True


def get_polygon_area(vertices_list):
    area = 0
    for i in range(len(vertices_list)):
        va = vertices_list[i]
        vb = vertices_list[(i + 1) % len(vertices_list)]
        width = vb[0] - va[0]
        height = (vb[1] + va[1]) / 2
        area += width * height
    return area


def reduce_vertices(vertices_list):
    if np.array_equal(vertices_list[0], vertices_list[-1]):
        vertices_list.pop(-1)
    n = len(vertices_list)
    index_list = list(range(0, n))
    delete_list = []
    for i in range(0, n):
        P0 = Point(vertices_list[i])
        P1 = Point(vertices_list[i - 1])
        P2 = Point(vertices_list[(i + 1) % n])
        if P0.is_collinear(Line(P1, P2)):
            delete_list.append(i)
    res_list = sorted(list(set(index_list) - set(delete_list)))
    reduced_vertices_list = list(itemgetter(*res_list)(vertices_list))
    return reduced_vertices_list


def get_coordinates(iterable, dimension=2):
    if not isinstance(iterable, np.ndarray):
        iterable = np.array(iterable)
    n = int(iterable.size / dimension)
    return np.split(iterable, n)


def triangulate_polygon(vertices_list, get_index=False):
    if vertices_list is None:
        return False
    vertices_list = get_coordinates(vertices_list)
    vertices_list = reduce_vertices(vertices_list)
    n = len(vertices_list)
    if n < 3 or n > 1024:
        raise Exception("Number of vertices exceed limits!")
    if get_polygon_area(vertices_list) > 0:
        counter = 1
    else:
        counter = -1
    index_list = list(range(0, n))
    triangles = []
    triangles_index = []
    while len(index_list) > 3:
        for i in range(0, n):
            a = index_list[i]
            b = index_list[i - 1]
            c = index_list[(i + 1) % n]
            va = vertices_list[a]
            vb = vertices_list[b]
            vc = vertices_list[c]
            vab = np.subtract(vb, va)
            vac = np.subtract(vc, va)
            if np.cross(vab, vac) * counter < 0:
                continue
            is_ear = True
            for j in range(0, len(vertices_list)):
                if j == a or j == b or j == c:
                    continue
                vp = vertices_list[j]
                if is_point_in_triangle(vp, vb, va, vc, counter):
                    is_ear = False
                    break
            if is_ear:
                triangles.append([vertices_list[b], vertices_list[a], vertices_list[c]])
                triangles_index.append([b, a, c])
                index_list.pop(i)
                break
    if get_index:
        return triangles_index
    else:
        return triangles


def get_wall_axes(vertices_list):
    triangles = triangulate_polygon(vertices_list)
    axes = []
    adjacent = []
    for i in range(len(triangles)):
        a_set = set(triangles[i])
        for j in range(len(triangles)):
            if i == j or [i, j] in adjacent:
                continue
            b_set = set(triangles[j])
            if len(a_set.intersection(b_set)) == 2:
                axes.append([triangles[i], triangles[j]])
                adjacent.append([j, i])
    return axes


if __name__ == '__main__':
    draftsman = CAD()
    assistant = Assistant()
    # draftsman.selection_set.Clear()
    # draftsman.selection_set.SelectOnScreen()
    # for number in range(draftsman.selection_set.Count):
    #     item = draftsman.selection_set.Item(number)
    #     coord = item.Coordinates
    #     # tri_coord = get_wall_axes(coord)
    #     # print(tri_coord)
    #     vertices = np.array(coord)
    #     total_vertex = int(vertices.size / 2)
    #     vertices = np.split(vertices, total_vertex)
    #     print(triangulate_polygon(vertices))

    # mid_points = []
    # for triangle in tri_coord:
    #     p0 = triangle[0]
    #     p1 = triangle[1]
    #     P0 = Point(p0[0], p0[1])
    #     P1 = Point(p1[0], p1[1])
    #     draftsman.draw_line(p0, p1)
    #     pm = P0.interpolate2point(P1, 0.5)
    #     mid_points.append(pm)
    # n = len(mid_points)
    # index_list = list(range(0, n))
    # lines = []
    # for i in range(2, len(index_list)):
    #     is_axis = False
    #     P0 = mid_points[index_list[i-2]]
    #     P1 = mid_points[index_list[i-1]]
    #     L0 = Line(P0, P1)
    #     for j in range(2, len(index_list)):
    #         P2 = mid_points[index_list[j]]
    #         if P2.is_collinear(L0):
    #             index_list.pop(j)
    #             is_axis = True
    #     if is_axis:
    #         index_list.pop(0)
    #         index_list.pop(1)
    #         lines.append(L0)
    # for line in lines:
    #     draftsman.draw_line(line.P0, line.P1, 'A-ACERO')

    # draftsman.select_all()
    # draftsman.move([0, 0, 0], [0, 5, 0])
    assistant.download_excel_beams_info()
    with open('beams_info.json') as jsonFile:
        beams_info = json.load(jsonFile)
    for name, info in beams_info.items():
        draftsman.draw_beam(info)
        draftsman.select_all()
        draftsman.move([0, 0, 0], [0, 5, 0])
    # beam_geo = [[0.25, 0.25, 0.25],
    #             [0.6, 0.8, 0.5],
    #             [5, 5, 6]]
    # beam_rest = [0, 1, 1.5, 0.4]
    # beam_re_info = [[[[0, 0, 1, 2, 5, 1, 0, 0, 0], [0, 0, 0, 2, 5, -1, 0, 0, 0]],
    #                  [[0, 1, 1, 2, 5, 1, 0, 0, 0], [0, 0, 0, 2, 5, -1, 0, 0, 0]],
    #                  [[0, 1, 1, 2, 5, 1, 0, 0, 0], [0, 0, 0, 2, 5, -1, 0, 0, 0]],
    #                  [[0, 1, 0, 2, 5, 1, 0, 0, 0], [0, 0, 0, 2, 5, -1, 0, 0, 0]]],
    #                 [[],
    #                  [],
    #                  [],
    #                  []]]
    draftsman.zoom_all()

# ANNOTATE
# AN1 = acad.model.AddDimAligned(PBase, PEnd, PAnnotateEnd)
# AN1 = acad.model.AddDimAngular(PBase, PEnd, PAnnotateEnd)

# CIRCLE, ARC & ELLIPSE
# C1 = acad.model.AddCircle(PBase, Radius)
# A1 = acad.model.AddArc(PBase, Radius, InitialAngle, FinalAngle)
# E1 = acad.model.AddEllipse(PBase1, PExt, RadiusRatio)

# MODIFY PARAMETERS
# acad.doc.GetVariable("PDMODE")
# acad.doc.SetVariable("PDMODE", 0)
