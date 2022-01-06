from FemtetClassConst import FemtetClassName as const
from win32com.client import Dispatch, constants
import sys
import FemtetClassConst
import importlib
FemtetClassConst = importlib.reload(FemtetClassConst)
# COMクラスのインスタンス作成、マクロ定数(constants)パッケージをインポート
# FemtetのCOMクラス名を定義したモジュールをインポート


def FemtetMain(model):
    # Femtet変数は別関数で利用したいのでglobal変数として定義する
    global Femtet
    Femtet = Dispatch(const.CFemtet)  # CFemtetクラスのインスタンス作成

    if Femtet.OpenNewProject() == False:
        print(Femtet.LastErrorMsg)
        sys.exit()

    # データベースの設定
    AnalysisSetUp()
    BodyAttributeSetUp()
    MaterialSetUp()
    BoundarySetUp()

    MakeModel(model)  # モデルの作成

    Femtet.Gaudi.MeshSize = model["mesh"]  # 標準メッシュサイズの設定

    Femtet.Gaudi.Mesh()  # メッシュの生成
    Femtet.Solve()  # 解析の実行

    return SamplingResult()  # 計算結果の表示


def AnalysisSetUp():

    global Femtet  # globalなFemtet変数を利用する
    Als = Femtet.Analysis  # 解析条件クラス

    # 解析条件共通(Common)
    Als.AnalysisType = constants.GAUSS_C

    # '------- 磁場(Gauss) -------
    Als.Gauss.b2ndEdgeElement = True
    Als.Gauss.bIncrementalInductance = False

    # '------- 電磁波(Hertz) -------
    Als.Hertz.b2ndEdgeElement = True

    # '------- 開放境界(Open) -------
    Als.Open.OpenMethod = constants.ABC_C
    Als.Open.ABCOrder = constants.ABC_2ND_C

    # '------- 調和解析(Harmonic) -------
    Als.Harmonic.FreqSweepType = constants.LINEAR_INTERVAL_C

    # '------- 高度な設定(HighLevel) -------
    Als.HighLevel.NonLTol = 1e-3
    Als.HighLevel.MemoryLimit = 32

    # '------- メッシュの設定(MeshProperty) -------
    Als.MeshProperty.bAutoAir = True
    Als.MeshProperty.AutoAirMeshSize = 69
    Als.MeshProperty.bChangePlane = True
    Als.MeshProperty.bMeshG2 = True
    Als.MeshProperty.bPeriodMesh = False

    # '------- 外部磁界(ExternalMagField) -------
    Als.ExternalMagField.FieldType = constants.EXTERNAL_B_C

    # '------- 結果インポート(Import) -------
    Als.Import.AnalysisModelName = "未選択"


# '////////////////////////////////////////////////////////////
# '    Body属性全体の設定
# '////////////////////////////////////////////////////////////
def BodyAttributeSetUp():

    # '------- Body属性の設定 -------
    BodyAttributeSetUp_ボディ属性_001()
    BodyAttributeSetUp_ボディ属性_003()
    BodyAttributeSetUp_ボディ属性_002()

    # '+++++++++++++++++++++++++++++++++++++++++
    # '++使用されていないBodyAttributeデータです
    # '++使用する際はコメントを外して下さい
    # '+++++++++++++++++++++++++++++++++++++++++
    # 'BodyAttributeSetUp_Air_Auto

# '////////////////////////////////////////////////////////////
# '    Body属性の設定 Body属性名：ボディ属性_002
# '////////////////////////////////////////////////////////////


def BodyAttributeSetUp_ボディ属性_001():
    global Femtet  # globalなFemtet変数を利用する
    BodyAttr = Femtet.BodyAttribute  # ボディ属性クラス

    # '------- Body属性の追加 -------
    BodyAttr.Add("ボディ属性_001")

    # '------- Body属性 Indexの設定 -------
    Index = BodyAttr.Ask("ボディ属性_001")

    # '------- シートボディの厚み or 2次元解析の奥行き(BodyThickness)/ワイヤーボディ幅(WireWidth) -------
    BodyAttr.Length(Index).bUseAnalysisThickness2D = True

    # '------- 波形(Waveform) -------
    BodyAttr.Current(Index).CurrentDirType = constants.COIL_NORMAL_INOUTFLOW_C

    # '------- 方向(Direction) -------
    BodyAttr.Direction(Index).SetVec(0, 0, -1)
    BodyAttr.Direction(Index).SetAxisVector(0, 0, 1)

    # '------- 初期速度(InitialVelocity) -------
    BodyAttr.InitialVelocity(Index).bAnalysisUse = True

    # '------- 流体(FluidBern) -------
    BodyAttr.FluidAttribute(Index).FlowCondition.bSpline = False

    # '------- 輻射(Emittivity) -------
    BodyAttr.ThermalSurface(Index).Emittivity.Eps = 0.8


# '////////////////////////////////////////////////////////////
# '    Body属性の設定 Body属性名：ボディ属性_001
# '////////////////////////////////////////////////////////////
def BodyAttributeSetUp_ボディ属性_003():
    global Femtet  # globalなFemtet変数を利用する
    BodyAttr = Femtet.BodyAttribute  # ボディ属性クラス

    # '------- Body属性の追加 -------
    BodyAttr.Add("ボディ属性_003")

    # '------- Body属性 Indexの設定 -------
    Index = BodyAttr.Ask("ボディ属性_003")

    # '------- シートボディの厚み or 2次元解析の奥行き(BodyThickness)/ワイヤーボディ幅(WireWidth) -------
    BodyAttr.Length(Index).bUseAnalysisThickness2D = True

    # '------- 方向(Direction) -------
    BodyAttr.Direction(Index).SetAxisVector(0, 0, 1)

    # '------- 初期速度(InitialVelocity) -------
    BodyAttr.InitialVelocity(Index).bAnalysisUse = True

    # '------- 流体(FluidBern) -------
    BodyAttr.FluidAttribute(Index).FlowCondition.bSpline = False

    # '------- 輻射(Emittivity) -------
    BodyAttr.ThermalSurface(Index).Emittivity.Eps = 0.8


def BodyAttributeSetUp_ボディ属性_002():
    global Femtet  # globalなFemtet変数を利用する
    BodyAttr = Femtet.BodyAttribute  # ボディ属性クラス

    # '------- Body属性の追加 -------
    BodyAttr.Add("ボディ属性_002")

    # '------- Body属性 Indexの設定 -------
    Index = BodyAttr.Ask("ボディ属性_002")

    # '------- シートボディの厚み or 2次元解析の奥行き(BodyThickness)/ワイヤーボディ幅(WireWidth) -------
    BodyAttr.Length(Index).bUseAnalysisThickness2D = True

    # '------- 方向(Direction) -------
    BodyAttr.Direction(Index).SetAxisVector(0, 0, 1)

    # '------- 初期速度(InitialVelocity) -------
    BodyAttr.InitialVelocity(Index).bAnalysisUse = True

    # '------- 流体(FluidBern) -------
    BodyAttr.FluidAttribute(Index).FlowCondition.bSpline = False

    # '------- 輻射(Emittivity) -------
    BodyAttr.ThermalSurface(Index).Emittivity.Eps = 0.8


# '////////////////////////////////////////////////////////////
# '    Body属性の設定 Body属性名：Air_Auto
# '////////////////////////////////////////////////////////////
def BodyAttributeSetUp_Air_Auto():
    global Femtet  # globalなFemtet変数を利用する
    BodyAttr = Femtet.BodyAttribute  # ボディ属性クラス

    # '------- Body属性の追加 -------
    BodyAttr.Add("Air_Auto")

    # '------- Body属性 Indexの設定 -------
    Index = BodyAttr.Ask("Air_Auto")

    # '------- シートボディの厚み or 2次元解析の奥行き(BodyThickness)/ワイヤーボディ幅(WireWidth) -------
    BodyAttr.Length(Index).bUseAnalysisThickness2D = True

    # '------- 解析領域(ActiveSolver) -------
    BodyAttr.ActiveSolver(Index).bWatt = False
    BodyAttr.ActiveSolver(Index).bGalileo = False

    # '------- 初期速度(InitialVelocity) -------
    BodyAttr.InitialVelocity(Index).bAnalysisUse = True

    # '------- ステータ/ロータ(StatorRotor) -------
    BodyAttr.StatorRotor(Index).State = constants.AIR_C

    # '------- 流体(FluidBern) -------
    BodyAttr.FluidAttribute(Index).FlowCondition.bSpline = False

    # '------- 輻射(Emittivity) -------
    BodyAttr.ThermalSurface(Index).Emittivity.Eps = 0.8


# '////////////////////////////////////////////////////////////
# '    Material全体の設定
# '////////////////////////////////////////////////////////////
def MaterialSetUp():

    # '------- Materialの設定 -------
    MaterialSetUp_000_ネオジム磁石()

    # '+++++++++++++++++++++++++++++++++++++++++
    # '++使用されていないMaterialデータです
    # '++使用する際はコメントを外して下さい
    # '+++++++++++++++++++++++++++++++++++++++++
    # 'MaterialSetUp_Air_Auto

# '////////////////////////////////////////////////////////////
# '    Materialの設定 Material名：002_フェライト磁石
# '////////////////////////////////////////////////////////////


def MaterialSetUp_000_ネオジム磁石():
    global Femtet  # globalなFemtet変数を利用する
    Mtl = Femtet.Material  # 材料定数クラス

    # '------- Materialの追加 -------
    Mtl.Add("000_ネオジム磁石")

    # '------- Material Indexの設定 -------
    Index = Mtl.Ask("000_ネオジム磁石")

    # '------- 透磁率(Permeability) -------
    Mtl.Permeability(
        Index).MagneticMaterialType = constants.MAGNETIC_PERMANENT_C
    Mtl.Permeability(Index).sMu = 1.05
    Mtl.Permeability(
        Index).BHExtrapolationType = constants.BH_GRADIENT_LASTTWOPOINT_C

    # '------- 密度(Density) -------
    Mtl.Density(Index).Dens = 7400

    # '------- 圧電定数(PiezoElectricity) -------
    Mtl.PiezoElectricity(Index).Set_mE(0, 0)
    Mtl.PiezoElectricity(Index).Set_mE(1, 0)
    Mtl.PiezoElectricity(Index).Set_mE(2, 0)
    Mtl.PiezoElectricity(Index).Set_mE(3, 0)
    Mtl.PiezoElectricity(Index).Set_mE(5, 0)
    Mtl.PiezoElectricity(Index).Set_mE(6, 0)
    Mtl.PiezoElectricity(Index).Set_mE(7, 0)
    Mtl.PiezoElectricity(Index).Set_mE(8, 0)
    Mtl.PiezoElectricity(Index).Set_mE(10, 0)
    Mtl.PiezoElectricity(Index).Set_mE(11, 0)
    Mtl.PiezoElectricity(Index).Set_mE(15, 0)
    Mtl.PiezoElectricity(Index).Set_mE(16, 0)
    Mtl.PiezoElectricity(Index).Set_mE(17, 0)

    # '------- 磁石(Magneto) -------
    Mtl.Magneto(Index).sM = 1.24

# '////////////////////////////////////////////////////////////
# '    Materialの設定 Material名：Air_Auto
# '////////////////////////////////////////////////////////////


def MaterialSetUp_Air_Auto():
    global Femtet  # globalなFemtet変数を利用する
    Mtl = Femtet.Material  # 材料定数クラス

    # '------- Materialの追加 -------
    Mtl.Add("Air_Auto")

    # '------- Material Indexの設定 -------
    Index = Mtl.Ask("Air_Auto")

    # '------- 誘電体(Permittivity) -------
    Mtl.Permittivity(Index).sEps = 1.000517


# '////////////////////////////////////////////////////////////
# '    Boundary全体の設定
# '////////////////////////////////////////////////////////////
def BoundarySetUp():

    # '------- Boundaryの設定 -------
    BoundarySetUp_RESERVED_default()


# '////////////////////////////////////////////////////////////
# '    Boundaryの設定 Boundary名：RESERVED_default (外部境界条件)
# '////////////////////////////////////////////////////////////
def BoundarySetUp_RESERVED_default():
    global Femtet  # globalなFemtet変数を利用する
    Bnd = Femtet.Boundary  # 境界条件クラス

    # '------- Boundaryの追加 -------
    Bnd.Add("RESERVED_default")

    # '------- Boundary Indexの設定 -------
    Index = Bnd.Ask("RESERVED_default")

    # '------- 電気(Electrical) -------
    Bnd.Electrical(Index).Condition = constants.ELECTRIC_WALL_C

    # '------- 熱(Thermal) -------
    Bnd.Thermal(Index).bConAuto = True
    Bnd.Thermal(Index).bSetRadioSetting = False

    # '------- 室温_環境温度(RoomTemp) -------
    Bnd.Thermal(Index).RoomTemp.TempType = constants.TEMP_AMBIENT_C

    # '------- 輻射(Emittivity) -------
    Bnd.Thermal(Index).Emittivity.Eps = 0.8

    # '------- 流体(FluidBern) -------
    Bnd.FluidBern(Index).bSpline = False

# '////////////////////////////////////////////////////////////
# '    IF関数
# '////////////////////////////////////////////////////////////
# Function F_IF(expression As Double, val_true As Double, val_false As Double) As Double
#    If expression Then
#        F_IF = val_true
#    Else
#        F_IF = val_false
#    End If

# End Function

# '////////////////////////////////////////////////////////////
# '    変数定義関数
# '////////////////////////////////////////////////////////////
# Sub InitVariables()


#    'VB上の変数の定義
#    pi = 3.14159265358979

#    'FemtetGUI上の変数の登録（既存モデルの変数制御等でのみ利用）

# End Sub

# '////////////////////////////////////////////////////////////
# '    モデル作成関数
# '////////////////////////////////////////////////////////////
def MakeModel(model):

    global Femtet  # globalなFemtet変数を利用する
    Gaudi = Femtet.Gaudi  # モデラークラス
    BodyList = []  # Body配列変数の定義
    BodyList2 = []

    # '------- モデルを描画させない設定 -------
    Femtet.RedrawMode = False

    # '------- CreateCylinder -------
    Point0 = Dispatch(const.CGaudiPoint)  # CGaudiPointクラスのインスタンス作成
    Point0.SetCoord(model["x"], model["y"], model["dis"] + model["Thick"])
    tmpBody = Gaudi.CreateCylinder(Point0, model["R"], 2)
    BodyList.append(tmpBody)  # リストに追加

    # '------- CreateCylinder -------
    Point1 = Dispatch(const.CGaudiPoint)  # CGaudiPointクラスのインスタンス作成
    Point1.SetCoord(model["rad"], 0, 0)
    tmpBody = Gaudi.CreateCylinder(Point1, model["R"], model["Thick"])
    BodyList2.append(tmpBody)  # リストに追加

    # '------- SetName -------
    BodyList[0].SetName("ボディ属性_001", "000_ネオジム磁石")

    # '------- SetName -------
    BodyList2[0].SetName("ボディ属性_002", "000_ネオジム磁石")

    # '------- RingCopy -------
    Point3 = Dispatch(const.CGaudiPoint)
    Vector0 = Dispatch(const.CGaudiVector)
    Point3.SetCoord(0, 0, 0)
    Vector0.SetCoord(0, 0, 1)
    ret = BodyList2[0].RingCopy_py(Point3, Vector0, 60, 5)
    BodyList2.extend(ret[1])
    # '------- VectorCopy -------
    Vector2 = Dispatch(const.CGaudiVector)
    Vector2.SetCoord(0, 0, -1)
    ret = Gaudi.MultiVectorCopy_py(
        BodyList2, Vector2, model["Thick"], model["mag_num"]-1)
    BodyList2.extend(ret[1])

    BodyList.extend(BodyList2)

    # '------- モデルを再描画します -------
    Femtet.Redraw()

# '////////////////////////////////////////////////////////////
# '    計算結果抽出関数
# '////////////////////////////////////////////////////////////


def SamplingResult():

    global Femtet  # globalなFemtet変数を利用する
    Gogh = Femtet.Gogh  # 解析結果クラス

    Femtet.SavePDT(Femtet.ResultFilePath + ".pdt", True)  # pdtファイルを保存します
    Femtet.OpenPDT(Femtet.ResultFilePath + ".pdt", True)  # pdtファイルを開きます

    # '------- フィールドの設定 -------
    Gogh.Gauss.Vector = constants.GAUSS_MAGNETIC_FLUX_DENSITY_C

    # 電磁力
    ret = Gogh.Gauss.GetMagForce_py("ボディ属性_001")
    print(ret)

    return ret
