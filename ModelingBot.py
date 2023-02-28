import os
import sys
import math
import win32com.client
import pythoncom
import comtypes.client
from typing import Union
from typing import Annotated


def clockwise_angle_and_distance(point, origin, ref_vec=[1, 0]):
    vector = [point[0] - origin[0], point[1] - origin[1]]
    len_vector = math.hypot(vector[0], vector[1])
    if len_vector == 0:
        return -math.pi, 0
    normalized = [vector[0] / len_vector, vector[1] / len_vector]
    dot_prod = normalized[0] * ref_vec[0] + normalized[1] * ref_vec[1]
    diff_prod = ref_vec[1] * normalized[0] - ref_vec[0] * normalized[1]
    angle = math.atan2(diff_prod, dot_prod)
    if angle < 0:
        return 2 * math.pi + angle, len_vector
    return angle, len_vector


def validate_coordinate_system(Dir, CSys):
    if Dir in [1, 2, 3] and CSys == "Local":
        return True
    elif Dir in [4, 5, 6, 7, 8, 9] and CSys != "Local":
        return True
    elif Dir in [10, 11] and CSys == "Global":
        return True
    else:
        return False


def convert_units(Units):
    if isinstance(Units, int):
        if 1 <= Units <= 16:
            return Units
        else:
            raise Exception("Invalid Units option ({0})".format(Units))
    if Units == "lb_in_F":
        Units = 1
    elif Units == "lb_ft_F":
        Units = 2
    elif Units == "kip_in_F":
        Units = 3
    elif Units == "kip_ft_F":
        Units = 4
    elif Units == "kN_mm_C":
        Units = 5
    elif Units == "kN_m_C":
        Units = 6
    elif Units == "kgf_mm_C":
        Units = 7
    elif Units == "kgf_m_C":
        Units = 8
    elif Units == "N_mm_C":
        Units = 9
    elif Units == "N_m_C":
        Units = 10
    elif Units == "tonf_mm_C":
        Units = 11
    elif Units == "tonf_m_C":
        Units = 12
    elif Units == "kN_cm_C":
        Units = 13
    elif Units == "kgf_cm_C":
        Units = 14
    elif Units == "N_cm_C":
        Units = 15
    elif Units == "tonf_cm_C":
        Units = 16
    else:
        raise Exception("Invalid Units option ({0})".format(Units))
    return Units


def convert_material_type(MatType):
    if isinstance(MatType, int):
        if 1 <= MatType <= 8:
            return MatType
        else:
            raise Exception("Invalid Material type ({0})".format(MatType))
    if MatType == "Steel":
        MatType = 1
    elif MatType == "Concrete":
        MatType = 2
    elif MatType == "NoDesign":
        MatType = 3
    elif MatType == "Aluminum":
        MatType = 4
    elif MatType == "ColdFormed":
        MatType = 5
    elif MatType == "Rebar":
        MatType = 6
    elif MatType == "Tendon":
        MatType = 7
    elif MatType == "Masonry":
        MatType = 8
    else:
        raise Exception("Invalid Material type ({0})".format(MatType))
    return MatType


def convert_load_pattern_type(MyType):
    if isinstance(MyType, int):
        if 1 <= MyType <= 12:
            return MyType
        else:
            raise Exception("Invalid Load Pattern type ({0})".format(MyType))
    if MyType == "Dead":
        MyType = 1
    elif MyType == "SuperDead":
        MyType = 2
    elif MyType == "Live":
        MyType = 3
    elif MyType == "ReduceLive":
        MyType = 4
    elif MyType == "Quake":
        MyType = 5
    elif MyType == "Wind":
        MyType = 6
    elif MyType == "Snow":
        MyType = 7
    elif MyType == "Other":
        MyType = 8
    elif MyType == "Move":
        MyType = 9
    elif MyType == "Temperature":
        MyType = 10
    elif MyType == "RoofLive":
        MyType = 11
    elif MyType == "Notional":
        MyType = 12
    else:
        raise Exception("Invalid Load Pattern type ({0})".format(MyType))
    return MyType


def convert_direction(Dir):
    if isinstance(Dir, int):
        if 1 <= Dir <= 11:
            return Dir
        else:
            raise Exception("Invalid Direction option ({0})".format(Dir))
    if Dir == "Local 1":
        Dir = 1
    elif Dir == "Local 2":
        Dir = 2
    elif Dir == "Local 3":
        Dir = 3
    elif Dir == "X":
        Dir = 4
    elif Dir == "Y":
        Dir = 5
    elif Dir == "Z":
        Dir = 6
    elif Dir == "Projected X":
        Dir = 7
    elif Dir == "Projected Y":
        Dir = 8
    elif Dir == "Projected Z":
        Dir = 9
    elif Dir == "Gravity":
        Dir = 10
    elif Dir == "Projected Gravity":
        Dir = 11
    else:
        raise Exception("Invalid Direction option ({0})".format(Dir))
    return Dir


def convert_combo_type(ComboType):
    if isinstance(ComboType, int):
        if 0 <= ComboType <= 4:
            return ComboType
        else:
            raise Exception("Invalid Combination type ({0})".format(ComboType))
    if ComboType == "Linear Additive":
        ComboType = 0
    elif ComboType == "Envelope":
        ComboType = 1
    elif ComboType == "Absolute Additive":
        ComboType = 2
    elif ComboType == "SRSS":
        ComboType = 3
    elif ComboType == "Range Additive":
        ComboType = 4
    else:
        raise Exception("Invalid Combination type ({0})".format(ComboType))
    return ComboType


def convert_item_type(ItemType):
    if isinstance(ItemType, int):
        if 0 <= ItemType <= 2:
            return ItemType
        else:
            raise Exception("Invalid ItemType option ({0})".format(ItemType))
    if ItemType == "Objects":
        ItemType = 1
    elif ItemType == "Group":
        ItemType = 2
    elif ItemType == "SelectedObjects":
        ItemType = 3
    else:
        raise Exception("Invalid ItemType option ({0})".format(ItemType))
    return ItemType


def convert_diaphragm_option(DiaphragmOption):
    if isinstance(DiaphragmOption, int):
        if 1 <= DiaphragmOption <= 3:
            return DiaphragmOption
        else:
            raise Exception("Invalid Diaphragm option ({0})".format(DiaphragmOption))
    if DiaphragmOption == "Disconnect":
        DiaphragmOption = 1
    elif DiaphragmOption == "From Shell Object":
        DiaphragmOption = 2
    elif DiaphragmOption == "Defined Diaphragm":
        DiaphragmOption = 3
    else:
        raise Exception("Invalid Diaphragm option ({0})".format(DiaphragmOption))
    return DiaphragmOption


def convert_slab_type(SlabType):
    if isinstance(SlabType, int):
        if SlabType == 2:
            raise Exception("Invalid Slab Type option ({0})".format(SlabType))
        elif 0 <= SlabType <= 6:
            return SlabType
        else:
            raise Exception("Invalid Slab Type option ({0})".format(SlabType))
    if SlabType == "Slab":
        SlabType = 0
    elif SlabType == "Drop":
        SlabType = 1
    elif SlabType == "Stiff":
        raise Exception("Invalid Slab Type option ({0})".format(SlabType))
    elif SlabType == "Ribbed":
        SlabType = 3
    elif SlabType == "Waffle":
        SlabType = 4
    elif SlabType == "Mat":
        SlabType = 5
    elif SlabType == "Footing":
        SlabType = 6
    else:
        raise Exception("Invalid Slab Type option ({0})".format(SlabType))
    return SlabType


def convert_shell_type(ShellType):
    if isinstance(ShellType, int):
        if 4 <= ShellType <= 5:
            raise Exception("Invalid Shell Type option ({0})".format(ShellType))
        elif 1 <= ShellType <= 6:
            return ShellType
        else:
            raise Exception("Invalid Shell Type option ({0})".format(ShellType))
    else:
        if ShellType == "Shell-Thin":
            ShellType = 1
        elif ShellType == "Shell-Thick":
            ShellType = 2
        elif ShellType == "Membrane":
            ShellType = 3
        elif ShellType == "Plate-Thin":
            raise Exception("Invalid Shell Type option ({0})".format(ShellType))
        elif ShellType == "Plate-Thick":
            raise Exception("Invalid Shell Type option ({0})".format(ShellType))
        elif ShellType == "Layered":
            ShellType = 6
        else:
            raise Exception("Invalid Shell Type option ({0})".format(ShellType))
    return ShellType


def convert_ribs_direction(RibsParallelTo):
    if isinstance(RibsParallelTo, int):
        if 1 <= RibsParallelTo <= 2:
            return RibsParallelTo
        else:
            raise Exception("Invalid Ribs direction ({0})".format(RibsParallelTo))
    if RibsParallelTo == "Local 1":
        RibsParallelTo = 1
    elif RibsParallelTo == "Local 2":
        RibsParallelTo = 2
    else:
        raise Exception("Invalid Ribs direction ({0})".format(RibsParallelTo))
    return RibsParallelTo


def convert_confine_type(ConfineType):
    if isinstance(ConfineType, int):
        if 1 <= ConfineType <= 2:
            return ConfineType
        else:
            raise Exception("Invalid Ribs direction ({0})".format(ConfineType))
    if ConfineType == "Ties":
        ConfineType = 1
    elif ConfineType == "Spiral":
        ConfineType = 2
    else:
        raise Exception("Invalid Ribs direction ({0})".format(ConfineType))
    return ConfineType


class SAP:
    def __init__(self, AttachToInstance=False, SpecifyPath=False):
        self.bar_size_names = []
        self.load_patterns_names = []
        self.load_cases_names = []
        self.load_combinations_names = []
        self.response_spectrum_functions = []
        self.AttachToInstance = AttachToInstance
        self.SpecifyPath = SpecifyPath
        self.ProgramPath = "C:\\Program Files\\Computers and Structures\\ETABS 19\\ETABS.exe"
        self.APIPath = "D:\\PROYECTOS\\PROGRAMACION\\CSIBotModels"
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
                self.Etabs = self.helper.GetObject("CSI.ETABS.API.ETABSObject")
            except (OSError, comtypes.COMError):
                print("No running instance of the program found or failed to attach.")
                sys.exit(-1)
        else:
            if self.SpecifyPath:
                try:
                    self.Etabs = self.helper.CreateObject(self.ProgramPath)
                except (OSError, comtypes.COMError):
                    print("Cannot start a new instance of the program from " + self.ProgramPath)
                    sys.exit(-1)
            else:
                try:
                    self.Etabs = self.helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject")
                except (OSError, comtypes.COMError):
                    print("Cannot start a new instance of the program.")
                    sys.exit(-1)
            self.Etabs.ApplicationStart()
        self.SapModel = self.Etabs.SapModel

    def initialize(self, Units: Union[str, int] = "tonf_m_C"):
        Units = convert_units(Units)
        return self.SapModel.InitializeNewModel(Units)

    def new_model(self, temp, *args):
        if args is None:
            args = [3, 3, 3, 4, 4, 5, 5]
        if temp == 1:
            ret = self.SapModel.File.NewBlank()
        elif temp == 2:
            ret = self.SapModel.File.NewGridOnly(args[0], args[1], args[2], args[3], args[4], args[5], args[6])
        elif temp == 3:
            ret = self.SapModel.File.NewSteelDeck(args[0], args[1], args[2], args[3], args[4], args[5], args[6])
        else:
            raise Exception("Not available template option")
        return ret

    def save_model(self):
        return self.SapModel.File.Save(self.ModelPath)

    def run_analysis(self):
        return self.SapModel.Analyze.RunAnalysis()

    def switch_units(self, Units: Union[str, int] = "tonf_m_C"):
        Units = convert_units(Units)
        return self.SapModel.SetPresentUnits(Units)

    def define_material(self, Name: str, MatType: Union[str, int], E: float, U: float, A: float, Temp=0):
        MatType = convert_material_type(MatType)
        self.SapModel.PropMaterial.SetMaterial(Name, MatType)
        return self.SapModel.PropMaterial.SetMPIsotropic(Name, E, U, A, Temp)

    def define_mass_source(self, IncludeElements: bool, IncludeAddedMass: bool, IncludeLoads: bool,
                           NumberLoads: int = 0, LoadPat: list[str] = None, SF: list[float] = None):
        if not IncludeLoads:
            NumberLoads = 0
            LoadPat = []
            SF = []
        elif NumberLoads == 0:
            raise Exception("Number of loads must be greater or equal to 1")
        else:
            if not isinstance(LoadPat, list):
                raise Exception("Argument LoadPat must be a list of NumberLoads elements")
            elif len(LoadPat) != NumberLoads:
                raise Exception("Number of LoadPat must be equal to number of loads")
            if not isinstance(SF, list):
                raise Exception("Argument SF must be a list of NumberLoads elements")
            elif len(SF) != NumberLoads:
                raise Exception("Number of SF must be equal to number of loads")
        for i in range(NumberLoads):
            if LoadPat[i] not in self.load_patterns_names:
                raise Exception("Load Pattern ({0}) not defined".format(LoadPat[i]))
        return self.SapModel.PropMaterial.SetMassSource_1(IncludeElements, IncludeAddedMass, IncludeLoads, NumberLoads,
                                                          LoadPat, SF)

    def define_rectangular_frame_property(self, Name: str, MatProp: str, T3: float, T2: float,
                                          Value: Annotated[list[float], 8] = None,
                                          A=1.0, V2=1.0, V3=1.0, T=1.0, M2=1.0, M3=1.0, Mm=1.0, Wm=1.0):
        self.SapModel.PropFrame.SetRectangle(Name, MatProp, T3, T2)
        if Value is None:
            Value = [A, V2, V3, T, M2, M3, Mm, Wm]
        return self.SapModel.PropFrame.SetModifiers(Name, Value)

    def define_property_beam_rebar(self, Name: str, MatPropLong: str, MatPropConfine: str, CoverTop: float,
                                   CoverBot: float, TopLeftArea: float, TopRightArea: float, BotLeftArea: float,
                                   BotRightArea: float):
        return self.SapModel.PropFrame.SetRebarBeam(Name, MatPropLong, MatPropConfine, CoverTop, CoverBot, TopLeftArea,
                                                    TopRightArea, BotLeftArea, BotRightArea)

    def define_rectangular_column_rebar(self, Name: str, MatPropLong: str, MatPropConfine: str, Cover: float,
                                        NumberR3Bars: int, NumberR2Bars: int, RebarSize: str, TieSize: str,
                                        TieSpacingLongit: float, Number2DirTieBars: int, Number3DirTieBars: int,
                                        ToBeDesigned: bool = False):
        Pattern = 1
        ConfineType = 1
        NumberCBars = 0
        if RebarSize not in self.bar_size_names:
            raise Exception("Not defined Rebar Size ({0})".format(RebarSize))
        if TieSize not in self.bar_size_names:
            raise Exception("Not defined Rebar Size ({0})".format(TieSize))
        return self.SapModel.PropFrame.SetRebarColumn(Name, MatPropLong, MatPropConfine, Pattern, ConfineType, Cover,
                                                      NumberCBars, NumberR3Bars, NumberR2Bars, RebarSize, TieSize,
                                                      TieSpacingLongit, Number2DirTieBars, Number3DirTieBars,
                                                      ToBeDesigned)

    def define_circular_column_rebar(self, Name: str, MatPropLong: str, MatPropConfine: str, Cover: float,
                                     NumberCBars: int, RebarSize: str, TieSize: str, TieSpacingLongit: float,
                                     ConfineType: Union[str, int] = "Ties", ToBeDesigned: bool = False):
        Pattern = 2
        ConfineType = convert_confine_type(ConfineType)
        NumberR3Bars = 0
        NumberR2Bars = 0
        Number2DirTieBars = 0
        Number3DirTieBars = 0
        if RebarSize not in self.bar_size_names:
            raise Exception("Not defined Rebar Size ({0})".format(RebarSize))
        if TieSize not in self.bar_size_names:
            raise Exception("Not defined Rebar Size ({0})".format(TieSize))
        return self.SapModel.PropFrame.SetRebarColumn(Name, MatPropLong, MatPropConfine, Pattern, ConfineType, Cover,
                                                      NumberCBars, NumberR3Bars, NumberR2Bars, RebarSize, TieSize,
                                                      TieSpacingLongit, Number2DirTieBars, Number3DirTieBars,
                                                      ToBeDesigned)

    def define_slab_shell_property(self, Name: str, ShellType: Union[str, int], MatProp: str, Thickness: float,
                                   SlabType: Union[str, int] = 0, Value: Annotated[list[float], 10] = None,
                                   F11=1, F22=1, F12=1, M11=1, M22=1, M12=1, V13=1, V23=1, Mm=1, Wm=1):
        if 3 <= SlabType <= 4:
            raise Exception("Use function ribbed or waffle for SlabType ({0})".format(SlabType))
        SlabType = convert_slab_type(SlabType)
        ShellType = convert_shell_type(ShellType)
        self.SapModel.PropArea.SetSlab(Name, SlabType, ShellType, MatProp, Thickness)
        if Value is None:
            Value = [F11, F22, F12, M11, M22, M12, V13, V23, Mm, Wm]
        return self.SapModel.PropArea.SetModifiers(Name, Value)

    def define_ribbed_shell_property(self, Name: str, ShellType: Union[str, int], MatProp: str, OverallDepth: float,
                                     SlabThickness: float, StemWidthTop: float, StemWidthBot: float, RibSpacing: float,
                                     RibsParallelTo: Union[str, int], Value: Annotated[list[float], 10] = None,
                                     F11=1, F22=1, F12=1, M11=1, M22=1, M12=1, V13=1, V23=1, Mm=1, Wm=1):
        SlabType = 3
        ShellType = convert_shell_type(ShellType)
        self.SapModel.PropArea.SetSlab(Name, SlabType, ShellType, MatProp, OverallDepth)
        RibsParallelTo = convert_ribs_direction(RibsParallelTo)
        self.SapModel.PropArea.SetSlabRibbed(Name, OverallDepth, SlabThickness, StemWidthTop, StemWidthBot,
                                             RibSpacing, RibsParallelTo)
        if Value is None:
            Value = [F11, F22, F12, M11, M22, M12, V13, V23, Mm, Wm]
        return self.SapModel.PropArea.SetModifiers(Name, Value)

    def define_waffle_shell_property(self, Name: str, ShellType: Union[str, int], MatProp: str, OverallDepth: float,
                                     SlabThickness: float, StemWidthTop: float, StemWidthBot: float,
                                     RibSpacingDir1: float, RibSpacingDir2: float,
                                     Value: Annotated[list[float], 10] = None,
                                     F11=1, F22=1, F12=1, M11=1, M22=1, M12=1, V13=1, V23=1, Mm=1, Wm=1):
        SlabType = 4
        ShellType = convert_shell_type(ShellType)
        self.SapModel.PropArea.SetSlab(Name, SlabType, ShellType, MatProp, OverallDepth)
        self.SapModel.PropArea.SetSlabWaffle(Name, OverallDepth, SlabThickness, StemWidthTop, StemWidthBot,
                                             RibSpacingDir1, RibSpacingDir2)
        if Value is None:
            Value = [F11, F22, F12, M11, M22, M12, V13, V23, Mm, Wm]
        return self.SapModel.PropArea.SetModifiers(Name, Value)

    def define_wall_shell_property(self, Name: str, ShellType: Union[str, int], MatProp: str, Thickness: float,
                                   Value: Annotated[list[float], 10] = None,
                                   F11=1, F22=1, F12=1, M11=1, M22=1, M12=1, V13=1, V23=1, Mm=1, Wm=1):
        WallPropType = 1
        ShellType = convert_shell_type(ShellType)
        self.SapModel.PropArea.SetWall(Name, WallPropType, ShellType, MatProp, Thickness)
        if Value is None:
            Value = [F11, F22, F12, M11, M22, M12, V13, V23, Mm, Wm]
        return self.SapModel.PropArea.SetModifiers(Name, Value)

    def define_load_pattern(self, Name: str, MyType: Union[str, int], SelfWTMultiplier=0, AddAnalysisCase=True):
        MyType = convert_load_pattern_type(MyType)
        self.load_patterns_names.append(Name)
        return self.SapModel.LoadPatterns.Add(Name, MyType, SelfWTMultiplier, AddAnalysisCase)

    def define_load_case_linear_static(self, Name: str, LoadName: Union[str, list[str]],
                                       SF: Union[float, list[float]] = None, InitialCase: str = "None"):
        self.SapModel.LoadCases.StaticLinear.SetCase(Name)
        self.load_cases_names.append(Name)
        self.SapModel.LoadCases.StaticLinear.SetInitialCase(Name, InitialCase)
        if isinstance(LoadName, list):
            NumberLoads = len(LoadName)
        else:
            LoadName = list(LoadName)
            NumberLoads = 1
        if SF is None:
            SF = [1.0] * NumberLoads
        elif isinstance(SF, float):
            SF = [SF] * NumberLoads
        elif len(SF) != NumberLoads:
            raise Exception("Number of SF must be equal to number of loads")
        LoadType = []
        for i in range(NumberLoads):
            if LoadName[i] in ["UX", "UY", "UZ", "RX", "RY", "RZ"]:
                LoadType.append("Accel")
            elif LoadName[i] in self.load_patterns_names:
                LoadType.append("Load")
            else:
                raise Exception("Load Pattern ({0}) not defined".format(LoadName[i]))
        return self.SapModel.LoadCases.StaticLinear.SetLoads(Name, NumberLoads, LoadType, LoadName, SF)

    def define_load_case_modal_eigen(self, Name: str = "Modal"):
        self.load_cases_names.append(Name)
        return self.SapModel.LoadCases.ModalEigen.SetCase(Name)

    def define_load_case_response_spectrum(self, Name: str, Eccen: float, LoadName: Union[str, list[str]],
                                           Func: Union[str, list[str]], SF: Union[float, list[float]],
                                           CSys: Union[str, list[str]] = None, Ang: Union[float, list[float]] = None,
                                           ModalCase: str = "Modal"):
        self.SapModel.LoadCases.ResponseSpectrum.SetCase(Name)
        self.load_cases_names.append(Name)
        self.SapModel.LoadCases.ResponseSpectrum.SetEccentricity(Name, Eccen)
        if ModalCase not in self.load_cases_names:
            raise Exception("Modal Case ({0}) not defined".format(ModalCase))
        self.SapModel.LoadCases.ResponseSpectrum.SetModalCase(Name, ModalCase)
        if isinstance(LoadName, list):
            NumberLoads = len(LoadName)
        else:
            LoadName = list(LoadName)
            NumberLoads = 1
        if isinstance(Func, str):
            Func = [Func] * NumberLoads
        elif len(Func) != NumberLoads:
            raise Exception("Number of Func must be equal to number of loads")
        if isinstance(SF, float):
            SF = [SF] * NumberLoads
        elif len(SF) != NumberLoads:
            raise Exception("Number of SF must be equal to number of loads")
        if CSys is None:
            CSys = ["Global"] * NumberLoads
        elif isinstance(CSys, str):
            CSys = [CSys] * NumberLoads
        elif len(CSys) != NumberLoads:
            raise Exception("Number of CSys must be equal to number of loads")
        if Ang is None:
            Ang = [0.0] * NumberLoads
        elif isinstance(Ang, float):
            Ang = [Ang] * NumberLoads
        elif len(Ang) != NumberLoads:
            raise Exception("Number of Ang must be equal to number of loads")
        for i in range(NumberLoads):
            if LoadName[i] not in ["U1", "U2", "U3", "R1", "R2", "R3"]:
                raise Exception("Direction({0}) not allowed. Please, use one of the following: ".format(LoadName[i]) +
                                "['U1', 'U2', 'U3', 'R1', 'R2', 'R3']")
            if Func[i] not in self.response_spectrum_functions:
                raise Exception("Response Spectrum Function ({0}) not defined".format(Func[i]))
        return self.SapModel.LoadCases.ResponseSpectrum.SetLoads(Name, NumberLoads, LoadName, Func, SF, CSys, Ang)

    def define_load_combination(self, Name: str, ComboType: Union[str, int], loadCNames: list[str],
                                loadCasesFactors: list[float] = None):
        ComboType = convert_load_pattern_type(ComboType)
        self.SapModel.RespCombo.Add(Name, ComboType)
        self.load_combinations_names.append(Name)
        for i in range(len(loadCNames)):
            CName = loadCNames[i]
            if CName in self.load_cases_names:
                CNameType = 0
            elif CName in self.load_combinations_names:
                CNameType = 1
            else:
                raise Exception("CName ({0}) not defined".format(CName))
            if loadCasesFactors is None:
                SF = 1
            else:
                SF = loadCasesFactors[i]
            self.SapModel.RespCombo.SetCaseList(Name, CNameType, CName, SF)

    def define_diaphragm(self, Name: str, SemiRigid: bool = False):
        return self.SapModel.Diaphragm.SetDiaphragm(Name, SemiRigid)

    def draw_frame(self, I_coord, J_coord, PropName="Default", UserName="", CSys="Global"):
        Name = ""
        XI, YI, ZI = I_coord[0], I_coord[1], I_coord[2]
        XJ, YJ, ZJ = J_coord[0], J_coord[1], J_coord[2]
        return self.SapModel.FrameObj.AddByCoord(XI, YI, ZI, XJ, YJ, ZJ, Name, PropName, UserName, CSys)

    def draw_frame_by_point(self, Point1: str, Point2: str, PropName="Default", UserName=""):
        Name = ""
        return self.SapModel.FrameObj.AddByPoint(Point1, Point2, Name, PropName, UserName)

    def delete_frame(self, Name: str, ItemType: Union[str, int] = 0):
        ItemType = convert_item_type(ItemType)
        return self.SapModel.FrameObj.Delete(Name, ItemType)

    def get_points(self, Name: str):
        Point1 = ""
        Point2 = ""
        return self.SapModel.FrameObj.GetPoints(Name, Point1, Point2)

    def get_releases(self, Name: str):
        II, JJ, StartValue, EndValue = [], [], [], []
        return self.SapModel.FrameObj.GetReleases(Name, II, JJ, StartValue, EndValue)

    def get_loads_distributed(self, Name: str, ItemType: Union[str, int] = 0):
        NumberItems, FrameName, LoadPat, MyType, CSys, Dir, RD1, RD2 = 0, [], [], [], [], [], [], []
        Dist1, Dist2, Val1, Val2 = [], [], [], []
        return self.SapModel.FrameObj.GetLoadDistributed(Name, NumberItems, FrameName, LoadPat, MyType, CSys, Dir,
                                                         RD1, RD2, Dist1, Dist2, Val1, Val2, ItemType)

    def assign_restraints(self, Name: str, Value: list[bool, bool, bool, bool, bool, bool] = None,
                          ItemType: Union[str, int] = 0, U1=False, U2=False, U3=False, R1=False, R2=False, R3=False):
        ItemType = convert_item_type(ItemType)
        if Value is None:
            Value = [U1, U2, U3, R1, R2, R3]
        return self.SapModel.PointObj.SetRestraint(Name, Value, ItemType)

    def assign_point_load(self, Name: str, LoadPat: str, Value: list[float, float, float, float, float, float] = None,
                          Replace=False, CSys="Global", ItemType: Union[str, int] = 0,
                          F1=0, F2=0, F3=0, M1=0, M2=0, M3=0):
        ItemType = convert_item_type(ItemType)
        if Value is None:
            Value = [F1, F2, F3, M1, M2, M3]
        return self.SapModel.PointObj.SetLoadForce(Name, LoadPat, Value, Replace, CSys, ItemType)

    def assign_frame_dist_load(self, Name: str, LoadPat: str, Dist1: float, Dist2: float, Val1: float, Val2: float,
                               Dir: Union[str, int] = "Gravity", MyType: int = 1, RelDist: bool = False,
                               Replace: bool = False, CSys: str = "Global", ItemType: Union[str, int] = 0):
        Dir = convert_direction(Dir)
        ItemType = convert_item_type(ItemType)
        if validate_coordinate_system(Dir, CSys):
            return self.SapModel.FrameObj.SetLoadDistributed(Name, LoadPat, MyType, Dir, Dist1, Dist2, Val1,
                                                             Val2, CSys, RelDist, Replace, ItemType)
        else:
            raise Exception("Not valid coordinate system for selected direction")

    def draw_shell(self, coordList, PropName="Default", UserName="", CSys="Global"):
        NumberPoints = len(coordList)
        X = []
        Y = []
        Z = []
        for coord in coordList:
            X.append(coord[0])
            Y.append(coord[1])
            Z.append(coord[2])
        Name = ""
        return self.SapModel.AreaObj.AddByCoord(NumberPoints, X, Y, Z, Name, PropName, UserName, CSys)

    def draw_shell_by_point(self, Point: list[str], PropName="Default", UserName=""):
        NumberPoints = len(Point)
        Name = ""
        return self.SapModel.AreaObj.AddByPoint(NumberPoints, Point, Name, PropName, UserName)

    def assign_shell_diaphragm(self, Name: str, DiaphragmName: str = "D1"):
        return self.SapModel.AreaObj.SetDiaphragm(Name, DiaphragmName)

    def assign_joint_diaphragm(self, Name: str, DiaphragmOption: str, DiaphragmName: str = "D1"):
        DiaphragmOption = convert_load_pattern_type(DiaphragmOption)
        return self.SapModel.FrameObj.SetDiaphragm(Name, DiaphragmOption, DiaphragmName)

    def assign_edge_constraint(self, Name: str, ConstraintExists: bool, ItemType: Union[str, int] = 0):
        ItemType = convert_item_type(ItemType)
        return self.SapModel.AreaObj.SetEdgeConstraint(Name, ConstraintExists, ItemType)

    def assign_shell_group(self, Name: str, GroupName: str, Remove: bool = False, ItemType: Union[str, int] = 0):
        ItemType = convert_item_type(ItemType)
        return self.SapModel.AreaObj.SetGroupAssign(Name, GroupName, Remove, ItemType)

    def assign_shell_uniform_load(self, Name: str, LoadPat: str, Value: float, Dir: Union[str, int] = "Gravity",
                                  Replace: bool = True, CSys: str = "Global", ItemType: Union[str, int] = 0):
        Dir = convert_direction(Dir)
        ItemType = convert_item_type(ItemType)
        if validate_coordinate_system(Dir, CSys):
            return self.SapModel.AreaObj.SetLoadUniform(Name, LoadPat, Value, Dir, Replace, CSys, ItemType)
        else:
            raise Exception("Not valid coordinate system for selected direction")

    def assign_shell_local_axes(self, Name: str, Ang: float, ItemType: Union[str, int] = 0):
        ItemType = convert_item_type(ItemType)
        return self.SapModel.AreaObj.SetLocalAxes(Name, Ang, ItemType)

    def delete_shell(self, Name: str, ItemType: Union[str, int] = 0):
        ItemType = convert_item_type(ItemType)
        return self.SapModel.AreaObj.Delete(Name, ItemType)

    def refresh_view(self, Window=0, Zoom=True):
        return self.SapModel.View.RefreshView(Window, Zoom)


if __name__ == "__main__":
    etabs = SAP()
    etabs.initialize(6)
    etabs.new_model(1)
    etabs.switch_units("kgf_cm_C")
    # # print(etabs.get_releases("1"))
    # # print(etabs.get_loads_distributed("1"))
    # # print(etabs.draw_frame([2, 2, 0], [3, 5, 0]))
    # # print(etabs.draw_frame([0, 1, 0], [2, 10, 0]))
    # # print(etabs.draw_shell([[0, 0, 0], [1, 0, 0], [1, 3, 0], [0, 3, 0]]))
    # # pts = [[2,3], [5,2],[4,1],[3.5,1],[1,2],[2,1],[3,1],[3,3],[4,3]]
    # # sort = sorted(pts, key=clockwise_angle_and_distance)
    # # print(sort)
    Bbeam_list = list(range(15, 65, 5))
    Hbeam_list = list(range(50, 105, 5)) + list(range(110, 210, 10))
    Fcbeam_list = [210, 280, 350]
    Bcol_list = list(range(15, 65, 5))
    Hcol_list = list(range(40, 165, 5))
    Fccol_list = [210, 280, 350, 420, 500]
    Es_list = [10, 15, 17, 20, 25]
    Ew_list = list(range(15, 55, 5))
    Ef_list = list(range(50, 110, 10))
    Fc_list = [175, 210, 280, 350, 420, 500]
    Econc_list = [198400, 217400, 251000, 280600, 307400, 335400]
    for index in range(len(Fc_list)):
        Fc = Fc_list[index]
        E = Econc_list[index]
        etabs.define_material("CONC" + str(Fc), "Concrete", E, 0.15, 0.0000099)
    for b_index in range(len(Bbeam_list)):
        B = Bbeam_list[b_index]
        for h_index in range(len(Hbeam_list)):
            H = Hbeam_list[h_index]
            for fc_index in range(len(Fcbeam_list)):
                Fc = Fcbeam_list[fc_index]
                etabs.define_rectangular_frame_property("V" + str(B) + "x" + str(H) + "-" + str(Fc) + "-AN",
                                                        "CONC" + str(Fc), H, B, T=0.0001)
    cols_created = []
    for b_index in range(len(Bcol_list)):
        B = Bcol_list[b_index]
        if 15 <= B <= 40:
            for h_index in range(len(Hcol_list)):
                H = Hcol_list[h_index]
                if B * 1000 + H in cols_created or H * 1000 + B in cols_created:
                    print("col no creada")
                    continue
                if H / B > 4:
                    continue
                for fc_index in range(len(Fccol_list)):
                    Fc = Fccol_list[fc_index]
                    etabs.define_rectangular_frame_property("C" + str(B) + "x" + str(H) + "-" + str(Fc),
                                                            "CONC" + str(Fc), H, B)
                    etabs.define_rectangular_frame_property("C" + str(H) + "x" + str(B) + "-" + str(Fc),
                                                            "CONC" + str(Fc), B, H)
                cols_created.append(B * 1000 + H)
                cols_created.append(H * 1000 + B)
        # elif 45 <= B <= 50:
        #     for h_index in range(len(Hcol_list)):
        #         H = Hcol_list[h_index]
        #         if B * 1000 + H in cols_created or H * 1000 + B in cols_created:
        #             print("col no creada")
        #             continue
        #         if H / B > 3:
        #             continue
        #         for fc_index in range(len(Fc_list)):
        #             Fc = Fc_list[fc_index]
        #             if Fc == 175:
        #                 continue
        #             etabs.define_rectangular_frame_property("C" + str(B) + "x" + str(H) + "-" + str(Fc),
        #                                                     "CONC" + str(Fc), H, B)
        #             etabs.define_rectangular_frame_property("C" + str(H) + "x" + str(B) + "-" + str(Fc),
        #                                                     "CONC" + str(Fc), B, H)
        #         cols_created.append(B * 1000 + H)
        #         cols_created.append(H * 1000 + B)
        # else:
        #     for h_index in range(len(Hcol_list)):
        #         H = Hcol_list[h_index]
        #         if B * 1000 + H in cols_created or H * 1000 + B in cols_created:
        #             print("col no creada")
        #             continue
        #         for fc_index in range(len(Fc_list)):
        #             Fc = Fc_list[fc_index]
        #             if Fc == 175:
        #                 continue
        #             etabs.define_rectangular_frame_property("C" + str(B) + "x" + str(H) + "-" + str(Fc),
        #                                                     "CONC" + str(Fc), H, B)
        #             etabs.define_rectangular_frame_property("C" + str(H) + "x" + str(B) + "-" + str(Fc),
        #                                                     "CONC" + str(Fc), B, H)
        #         cols_created.append(B * 1000 + H)
        #         cols_created.append(H * 1000 + B)
    # print(cols_created)
