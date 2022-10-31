import os
import sys
import math
import win32com.client
import pythoncom
import comtypes.client


def clockwise_angle_and_distance(point, origin, ref_vec=[1, 0]):
    vector = [point[0]-origin[0], point[1]-origin[1]]
    len_vector = math.hypot(vector[0], vector[1])
    if len_vector == 0:
        return -math.pi, 0
    normalized = [vector[0]/len_vector, vector[1]/len_vector]
    dot_prod = normalized[0]*ref_vec[0] + normalized[1]*ref_vec[1]
    diff_prod = ref_vec[1]*normalized[0] - ref_vec[0]*normalized[1]
    angle = math.atan2(diff_prod, dot_prod)
    if angle < 0:
        return 2*math.pi+angle, len_vector
    return angle, len_vector


class CAD:
    def __init__(self):
        self.acad = win32com.client.Dispatch("AutoCAD.Application")
        self.acadModel = self.acad.ActiveDocument.ModelSpace


class SAP:
    def __init__(self):
        self.AttachToInstance = True
        self.SpecifyPath = True
        self.ProgramPath = "C:\Program Files\Computers and Structures\ETABS 19"
        self.APIPath = "D:\PROYECTOS\PROGRAMACION\CSIBotModels"
        if not os.path.exists(self.APIPath):
            try:
                os.makedirs(self.APIPath)
            except OSError:
                pass
        self.ModelPath = self.APIPath + os.sep + 'API_1-001.edb'
        self.helper = comtypes.client.CreateObject('ETABSv1.Helper')
        self.helper = self.helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
        if self.AttachToInstance:
            try:
                self.etabs = self.helper.GetObject("CSI.ETABS.API.ETABSObject")
            except (OSError, comtypes.COMError):
                print("No running instance of the program found or failed to attach.")
                sys.exit(-1)
        else:
            if self.SpecifyPath:
                try:
                    self.etabs = self.helper.CreateObject(self.ProgramPath)
                except (OSError, comtypes.COMError):
                    print("Cannot start a new instance of the program from " + self.ProgramPath)
                    sys.exit(-1)
            else:
                try:
                    self.etabs = self.helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject")
                except (OSError, comtypes.COMError):
                    print("Cannot start a new instance of the program.")
                    sys.exit(-1)
            self.etabs.ApplicationStart()
        self.sapModel = self.etabs.SapModel

    def initialize(self, cUnits=12):
        self.sapModel.InitializeNewModel(cUnits)

    def new_model(self, temp, *args):
        if temp == 1:
            ret = self.sapModel.File.NewBlank()
        elif temp == 2:
            ret = self.sapModel.File.NewGridOnly(args[0], args[1], args[2], args[3], args[4], args[5], args[6])
        elif temp == 3:
            ret = self.sapModel.File.NewSteelDeck(args[0], args[1], args[2], args[3], args[4], args[5], args[6])
        else:
            ret = 1
        return ret

    def save_model(self):
        return self.sapModel.File.Save(self.ModelPath)

    def run_analysis(self):
        return self.sapModel.Analyze.RunAnalysis()

    def switch_units(self, tUnits="tonf_m_C"):
        if tUnits == "lb_in_F":
            eUnits = 1
        elif tUnits == "lb_ft_F":
            eUnits = 2
        elif tUnits == "kip_in_F":
            eUnits = 3
        elif tUnits == "kip_ft_F":
            eUnits = 4
        elif tUnits == "kN_mm_C":
            eUnits = 5
        elif tUnits == "kN_m_C":
            eUnits = 6
        elif tUnits == "kgf_mm_C":
            eUnits = 7
        elif tUnits == "kgf_m_C":
            eUnits = 8
        elif tUnits == "N_mm_C":
            eUnits = 9
        elif tUnits == "N_m_C":
            eUnits = 10
        elif tUnits == "tonf_mm_C":
            eUnits = 11
        elif tUnits == "tonf_m_C":
            eUnits = 12
        elif tUnits == "kN_cm_C":
            eUnits = 13
        elif tUnits == "kgf_cm_C":
            eUnits = 14
        elif tUnits == "N_cm_C":
            eUnits = 15
        elif tUnits == "tonf_cm_C":
            eUnits = 16
        else:
            return
        return self.sapModel.SetPresentUnits(eUnits)

    def define_material(self, name, tMatType, E, U, A, Temp=0):
        if tMatType == "STEEL":
            eMatType = 1
        elif tMatType == "CONCRETE":
            eMatType = 2
        elif tMatType == "NO_DESIGN":
            eMatType = 3
        elif tMatType == "ALUMINUM":
            eMatType = 4
        elif tMatType == "COLD_FORMED":
            eMatType = 5
        elif tMatType == "REBAR":
            eMatType = 6
        elif tMatType == "TENDON":
            eMatType = 7
        elif tMatType == "MASONRY":
            eMatType = 8
        else:
            return
        self.sapModel.PropMaterial.SetMaterial(name, eMatType)
        return self.sapModel.PropMaterial.SetMPIsotropic(name, E, U, A, Temp)

    def define_rectangular_frame_section(self, name, matName, B, H, A=1, V2=1, V3=1, T=1, M2=1, M3=1, Mm=1, Wm=1):
        self.sapModel.PropFrame.SetRectangle(name, matName, B, H)
        modifiers = [A, V2, V3, T, M2, M3, Mm, Wm]
        return self.sapModel.PropFrame.SetModifiers(name, modifiers)

    def define_load_pattern(self, name, tType, SW_multiplier=0, addCase=True):
        if tType == "DEAD":
            eType = 1
        elif tType == "SUPER_DEAD":
            eType = 2
        elif tType == "LIVE":
            eType = 3
        elif tType == "REDUCE_LIVE":
            eType = 4
        elif tType == "QUAKE":
            eType = 5
        elif tType == "WIND":
            eType = 6
        elif tType == "SNOW":
            eType = 7
        elif tType == "OTHER":
            eType = 8
        elif tType == "MOVE":
            eType = 9
        elif tType == "TEMPERATURE":
            eType = 10
        elif tType == "ROOF_LIVE":
            eType = 11
        elif tType == "NOTIONAL":
            eType = 12
        else:
            return
        return self.sapModel.LoadPatterns.Add(name, eType, SW_multiplier, addCase)

    def draw_frame(self, iCoord, fCoord, propName="Default", userName="LCM", CSys="Global"):
        frameName = " "
        return self.sapModel.FrameObj.AddByCoord(iCoord[0], iCoord[1], iCoord[2], fCoord[0], fCoord[1], fCoord[2]
                                                 , frameName, propName, userName, CSys)

    def get_points(self, frameName):
        pointName1 = " "
        pointName2 = " "
        return self.sapModel.FrameObj.GetPoints(frameName, pointName1, pointName2)

    def assign_restraints(self, pointName, U1=False, U2=False, U3=False, R1=False, R2=False, R3=False, itemType=0):
        restraints = [U1, U2, U3, R1, R2, R3]
        return self.sapModel.PointObj.SetRestraint(pointName, restraints, itemType)

    def assign_point_load(self, pointName, patternName, F1=0, F2=0, F3=0, M1=0, M2=0, M3=0, replace=False, CSys="Global"
                          , itemType=0):
        forces = [F1, F2, F3, M1, M2, M3]
        return self.sapModel.PointObj.SetLoadForce(pointName, patternName, forces, replace, CSys, itemType)

    def assign_frame_dist_load(self, frameName, patternName, dist1, dist2, val1, val2, tDir="Gravity", eType=1,
                               relDist=False, replace=False, CSys="Global", itemType=0):
        if tDir == "Local 1":
            eDir = 1
            CSys = "Local"
        elif tDir == "Local 2":
            eDir = 2
            CSys = "Local"
        elif tDir == "Local 3":
            eDir = 3
            CSys = "Local"
        elif tDir == "X":
            eDir = 4
        elif tDir == "Y":
            eDir = 5
        elif tDir == "Z":
            eDir = 6
        elif tDir == "Projected X":
            eDir = 7
        elif tDir == "Projected Y":
            eDir = 8
        elif tDir == "Projected Z":
            eDir = 9
        elif tDir == "Gravity":
            eDir = 10
            CSys = "Global"
        elif tDir == "Projected Gravity":
            eDir = 11
            CSys = "Global"
        else:
            return
        return self.sapModel.FrameObj.SetLoadDistributed(frameName, patternName, eType, eDir, dist1, dist2, val1, val2,
                                                         CSys, relDist, replace, itemType)

    def draw_area(self, coordList, propName="Default", userName="LCM", CSys="Global"):
        numPoints = len(coordList)
        XList = []
        YList = []
        ZList = []
        for coord in coordList:
            XList.append(coord[0])
            YList.append(coord[1])
            ZList.append(coord[2])
        areaName = " "
        return self.sapModel.AreaObj.AddByCoord(numPoints, XList, YList, ZList, areaName, propName, userName, CSys)

    def draw_area_by_point(self, pointNamesList, propName="Deafault", userName="LCM"):
        numPoints = len(pointNamesList)
        areaName = " "
        return self.sapModel.AreaObj.AddByPoint(numPoints, pointNamesList, areaName, propName, userName)

    def refresh_view(self, num=0, zoom=True):
        return self.sapModel.View.RefreshView(num, zoom)


if __name__ == "__main__":
    pts = [[2,3], [5,2],[4,1],[3.5,1],[1,2],[2,1],[3,1],[3,3],[4,3]]
    sort = sorted(pts, key=clockwise_angle_and_distance)
    print(sort)