import win32com.client
import pythoncom
# import comtypes.client
import math
from fractions import Fraction
import numpy as np
import time
import AssistantBot
from operator import itemgetter


def aDouble(xyz):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, xyz)


def aDispatch(vObject):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, vObject)


class Point:
    def __init__(self, x, y=0.0, z=0.0):
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

    def distance2point(self, P0):
        return math.sqrt((self.x - P0.x) ** 2 +
                         (self.y - P0.y) ** 2 +
                         (self.z - P0.z) ** 2)

    def distance2line(self, L0):
        x, y, z = self.x, self.y, self.z
        A, B, C = L0.A, L0.B, L0.C
        return math.fabs(A * x + B * y + C) / math.sqrt(A ** 2 + B ** 2)

    def projection2line(self, L0):
        x, y, z = self.x, self.y, self.z
        m, b = L0.m, L0.b
        m2 = -1 / m
        b2 = y - m2 * x
        int_x = -(b - b2) / (m - m2)
        int_y = m2 * int_x + b2
        return Point(int_x, int_y, z)

    def rotation(self, c, angle):
        x, y, z = self.x, self.y, self.z
        cx, cy = c
        s = math.sin(angle)
        c = math.cos(angle)
        x = x - cx
        y = y - cy
        x_new = x * c - y * s
        y_new = x * s + y * c
        x = x_new + cx
        y = y_new + cy
        return Point(x, y, z)

    def interpolate2point(self, P0, alpha):
        x, y, z = self.x, self.y, self.z
        x0, y0, z0 = P0.x, P0.y, P0.z
        return Point(x0 * alpha + x * (1 - alpha), y0 * alpha + y * (1 - alpha), z0 * alpha + z * (1 - alpha))

    def is_collinear(self, L0):
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

    def line2general(self):
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

    def intersect2line(self, L0):
        A, B, C = self.A, self.B, -self.C
        A0, B0, C0 = L0.A, L0.B, -L0.C
        if A * B0 - A0 * B == 0:
            raise Exception("Lines are parallel. There no intersection")
        else:
            return Point((C * B0 - C0 * B) / (A * B0 - A0 * B), (A * C0 - A0 * C) / (A * B0 - A0 * B))

    def mid_point(self):
        return self.P0.interpolate2point(self.P1, 0.5)

    def is_parallel(self, L0):
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

    def is_same(self, L0):
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


def get_dev_length(D, tie_case):
    lengths = AssistantBot.get_variable_value("DEV_LENGTHS")
    if tie_case != 2:
        # if case == 0:
        #     ld = 0
        # else:
        ld = lengths.get(str(tie_case) + str(D))
        return ld
    else:
        [ld, tie_long] = lengths.get(str(tie_case) + str(D))
        return [ld, tie_long]


def determine_tie_case(bar_case, bar_restriction, side):
    if side == 1:
        case = 0
        momentum_sign = get_envelope_momentum_signs(bar_case)[0]
    elif side == -1:
        case = 1
        momentum_sign = get_envelope_momentum_signs(bar_case)[1]

    return case


def get_envelope_momentum_signs(bar_case):
    AssistantBot.get_variable_value("ACTUAL_BEAM_FORCES_INFO").get(str(bar_case))
    return sign


class CAD:
    def __init__(self, file_name='Drawing1.dwg'):
        self.acad = win32com.client.Dispatch("AutoCAD.Application")
        # self.acad = Autocad(create_if_not_exists=True, visible=True)
        try:
            self.acad.ActiveDocument
        except Exception:
            self.acad.Documents.Add("Drawing1.dwg")
            time.sleep(3)
        documents = []
        for doc in self.acad.Documents:
            documents.append(doc)
            if self.acad.ActiveDocument.Name != file_name:
                self.acad.ActiveDocument = documents[-1]
            else:
                break
        self.acadDoc = self.acad.ActiveDocument
        self.acadModel = self.acadDoc.ModelSpace
        self.objects_list = []
        self.selection_set = self.acadDoc.ActiveSelectionSet
        self.selected_objects = []
        self.layers_list = []
        self.create_new_layer('A-TRAZO', 7)
        self.create_new_layer('A-ACERO', 4)
        self.create_new_layer('A-ESTRIBOS', 1)
        self.create_new_layer('A-TEXTOS', 3)
        self.create_new_layer('A-COTAS', 1)
        self.acadDoc.ActiveLayer = self.layers_list[0]
        # self.create_new_dim_style()

    def create_new_dim_style(self, name="1-100"):
        new_style = self.acad.ActiveDocument.DimStyles.Add(name)
        self.acadDoc.ActiveDimStyle = new_style
        self.acadDoc.SetVariable("DIMALTD", 2)

    def create_new_layer(self, name, color_num=1, line_type='Continuous', line_weight='Default'):
        new_layer = self.acadDoc.Layers.Add(name)
        new_layer.color = color_num
        try:
            self.acadDoc.Linetypes.Load(line_type, 'acadiso.lin')
        except Exception:
            pass
        finally:
            new_layer.Linetype = line_type
        if line_weight != 'Default':
            new_layer.Lineweight = line_weight
        self.layers_list.append(new_layer)

    def draw_beam(self, beam_geometry, beam_restraints, beam_reinforcement_info):
        h_max = max(beam_geometry[1])
        lt = sum(beam_geometry[2]) + sum(beam_restraints)
        ap = [0.5 * beam_restraints[0]]  # left axis position
        span_number = len(beam_restraints) - 1
        if beam_restraints[0] != 0:
            self.draw_line([0, 0.4, 0], [0, -h_max - 0.4, 0])
        if beam_restraints[-1] != 0:
            self.draw_line([lt, 0.4, 0], [lt, -h_max - 0.4, 0])
        for i in range(span_number):
            w = beam_geometry[0][i]  # width
            h = beam_geometry[1][i]  # height
            fl = beam_geometry[2][i]  # free length
            slr = beam_restraints[i] * 0.5  # span_left_restraint
            srr = beam_restraints[i + 1] * 0.5  # span_right_restraint
            self.draw_line([ap[i] + slr, 0, 0], [ap[i] + slr + fl, 0, 0])
            self.draw_line([ap[i] + slr, 0, 0], [ap[i] + slr, 0.4, 0]) if slr != 0 else \
                self.draw_line([ap[i] + slr, 0, 0], [ap[i] + slr, -0.5 * h, 0])
            self.draw_line([ap[i] + slr + fl, 0, 0], [ap[i] + slr + fl, 0.4, 0]) if srr != 0 else \
                self.draw_line([ap[i] + slr + fl, 0, 0], [ap[i] + slr + fl, -0.5 * h, 0])
            self.select_last(3)
            self.mirror([ap[i] + slr, -0.5 * h, 0], [ap[i] + slr + fl, -0.5 * h, 0])
            if h < h_max and slr != 0:
                self.draw_line([ap[i] + slr, -h_max - 0.4, 0], [ap[i] + slr, -h - 0.4, 0])
            if h < h_max and srr != 0:
                self.draw_line([ap[i] + slr + fl, -h_max - 0.4, 0], [ap[i] + slr + fl, -h - 0.4, 0])
            ap.append(ap[i] + slr + fl + srr)
        for i in range(len(beam_restraints)):
            if beam_restraints[i] != 0:
                self.draw_concrete_extension([ap[i] - beam_restraints[i] * 0.5, 0.4, 0],
                                             [ap[i] + beam_restraints[i] * 0.5, 0.4, 0])
                self.select_last(5)
                self.copy([0, 0.4, 0], [0, -h_max - 0.4, 0])
        for i in range(span_number):
            for bar_data in beam_reinforcement_info[0][i]:
                # bar_data = bar_data + [beam_restraints[i], beam_restraints[i + 1]]
                if bar_data[0] == 0:
                    if bar_data[7] == 1:
                        bar_data[5] -= beam_restraints[i] * 0.5
                    if bar_data[8] == 1:
                        bar_data[6] += beam_restraints[i + 1] * 0.5
                elif bar_data[0] == 1:
                    if bar_data[7] == 1:
                        bar_data[5] -= beam_restraints[i] * 0.5
                elif bar_data[0] == 2:
                    pass
                elif bar_data[0] == 3:
                    if bar_data[8] == 1:
                        bar_data[6] += beam_restraints[i + 1] * 0.5
                self.draw_beam_longitudinal_bar([ap[i] + beam_restraints[i] * 0.5, beam_geometry[1][i] * 0.5], bar_data)
            length_over = beam_geometry[2][i]
            last_position = 0
            if len(beam_reinforcement_info[1][i]) == 1 and beam_reinforcement_info[1][i][0][0] == 0:
                stirrups_offset = beam_reinforcement_info[1][i][0][1]
                stirrups_spacing = beam_reinforcement_info[1][i][0][2]
                stirrups_number = math.floor((length_over - stirrups_offset) / (2 * stirrups_spacing)) + 1
                beam_reinforcement_info[1][i] = [[1, stirrups_offset, stirrups_spacing, stirrups_number],
                                                 [2, stirrups_offset, stirrups_spacing, stirrups_number]]
            for stirrup_data in beam_reinforcement_info[1][i]:
                stirrups_spacing = stirrup_data[2]
                stirrups_number = stirrup_data[3]
                if stirrup_data[0] == 1:
                    start_point = ap[i] + beam_restraints[i] * 0.5 + stirrup_data[1]
                    last_position = stirrup_data[1] + stirrups_spacing * (stirrups_number - 1)
                    length_over -= last_position
                elif stirrup_data[0] == 2:
                    start_point = ap[i + 1] - beam_restraints[i + 1] * 0.5 - + stirrup_data[1]
                    length_over -= stirrup_data[1] + stirrups_spacing * (stirrups_number - 1)
                    stirrups_spacing = -stirrup_data[2]
                else:
                    length_over = round(length_over, 4)
                    aux = round((length_over - 2 * stirrup_data[1]) % stirrups_spacing, 4)
                    if aux == stirrups_spacing:
                        aux = 0
                        stirrups_number = (length_over - 2 * stirrup_data[1]) / stirrups_spacing + 1
                    else:
                        stirrups_number = math.ceil((length_over - 2 * stirrup_data[1]) / stirrups_spacing)
                    start_point = ap[i] + beam_restraints[i] * 0.5 + round(
                        last_position + stirrup_data[1] + 0.5 * aux, 4)
                self.draw_line(Point(start_point, -0.06, 0), Point(start_point, - beam_geometry[1][i] + 0.06, 0),
                               'A-ESTRIBOS')
                self.select_last(1)
                self.array(1, stirrups_number, 0, stirrups_spacing)

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

    def draw_line(self, P0, P1, layer='A-TRAZO'):
        if not isinstance(P0, Point):
            P0 = Point(P0)
        if not isinstance(P1, Point):
            P1 = Point(P1)
        L1 = self.acadModel.AddLine(P0.APoint, P1.APoint)
        L1.layer = layer
        self.objects_list.append(L1)

    def draw_polyline(self, points, layer='A-TRAZO'):
        points = aDouble(points)
        PL1 = self.acadModel.AddPolyline(points)
        PL1.layer = layer
        self.objects_list.append(PL1)

    def draw_text(self, P0, TSize=0.05, MText=False, BoxWidth=0, layer='A-TRAZO'):
        if not isinstance(P0, Point):
            P0 = Point(P0)
        if MText:
            T1 = self.acadModel.AddMText(P0.APoint, BoxWidth, "TEXT1")
        else:
            T1 = self.acadModel.AddText("TEXT1", P0.APoint, TSize)
        T1.layer = layer
        self.objects_list.append(T1)

    def draw_concrete_extension(self, P0, P1, fixed_height=0.2, ratio=0.0):
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
        self.draw_line(P0, P0p)
        self.draw_line(P0p, P2t)
        self.draw_line(P2t, P2b)
        self.draw_line(P2b, P1p)
        self.draw_line(P1p, P1)

    def draw_beam_longitudinal_bar(self, ip, bar_data, bar_restrictions=None, tie_case=-1):
        if bar_restrictions is None:
            bar_restrictions = [0, 0]
        bar_case = bar_data[0]
        N = bar_data[1]
        D = bar_data[2]
        side = bar_data[3]
        order = bar_data[4]
        left_cut = bar_data[5]
        right_cut = bar_data[6]
        left_con = bar_data[7]
        right_con = bar_data[8]
        edge_offset = 0.06 + 0.04 * order
        self.draw_line(Point(ip[0] + left_cut, -ip[1] + (ip[1] - edge_offset) * side),
                       Point(ip[0] + right_cut, -ip[1] + (ip[1] - edge_offset) * side), 'A-ACERO')
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

    def draw_tie_long_bar(self, P0, ld, case=0, side=1, draw_to=-1, tie_long=0):
        P1 = Point(P0.x + ld * draw_to, P0.y)
        self.draw_line(P0, P1, 'A-ACERO')
        if case == 1:
            P2 = Point(P1.x, P1.y - tie_long * side)
            self.draw_line(P1, P2, 'A-ACERO')

    def select_last(self, num_objects, selection_offset=0):
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
    res_list = sorted(list(set(index_list)-set(delete_list)))
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
    draftsman = CAD('Drawing1.dwg')
    draftsman.selection_set.Clear()
    draftsman.selection_set.SelectOnScreen()
    for number in range(draftsman.selection_set.Count):
        item = draftsman.selection_set.Item(number)
        coord = item.Coordinates
        # tri_coord = get_wall_axes(coord)
        # print(tri_coord)
        vertices = np.array(coord)
        total_vertex = int(vertices.size / 2)
        vertices = np.split(vertices, total_vertex)
        print(triangulate_polygon(vertices))

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
    # draftsman.draw_beam(beam_geo, beam_rest, beam_re_info)
    # draftsman.zoom_all()

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
