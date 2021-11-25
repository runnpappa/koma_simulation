from fem_excel import *
from fem_MagSize import model_size
from fem_Main_2 import FemtetMain as Run
from fem_Main_3 import FemtetMain as Run3


def rad_MagNum(rad_from, rad_to, MagNum_from, MagNum_to, RAD=True, MAG_NUM=True, HEIGHT=True):
    model = model_size()
    model["mag_num"] = MagNum_from
    if RAD != True or MAG_NUM != True or HEIGHT != True:
        Count.BackUpCount = False
    if MAG_NUM != True:
        model["mag_num"] = MAG_NUM
        MAG_NUM = True
    while model["mag_num"] <= MagNum_to:
        model["rad"] = rad_from
        if RAD != True:
            model["rad"] = RAD
            RAD = True
        while model["rad"] <= rad_to:
            if HEIGHT == True or HEIGHT == 40:
                E2_org("dis")
            model["dis"] = 40
            if HEIGHT != True:
                model["dis"] = HEIGHT
                HEIGHT = True
            while model["dis"] <= 70:
                E2main(Run3(model), model, "dis")
                model["dis"] += 5
            E2_comment("dis", "rad"+str(model["rad"])+"mm")
            E2_comment("dis", "mag_num"+str(model["mag_num"]))
            model["rad"] += 1
        model["mag_num"] += 1
        if model["mag_num"]>=10:
            model["mesh"]=0.6


def var_dis3(var, num, dis_from, dis_to, sub=False):
    wk.book = "koma_sim3_py.xlsx"
    for i in num:
        model = model_size()
        model[var] = i
        model["dis"] = dis_from
        E2_org("dis")
        while model["dis"] <= dis_to:
            E2main(Run3(model), model, "dis")
            if sub == True:
                model["dis"] += 10
            else:
                model["dis"] += 5
        E2_comment("dis", var+str(i))
    wk.book = "koma_sim2_py.xlsx"


def var_dis(var, num, From, To, sub=False):
    for i in num:
        model = model_size()
        model[var] = i
        model["dis"] = From
        E2_org("dis")
        while model["dis"] <= To:
            E2main(Run(model), model, "dis")
            if sub == True:
                model["dis"] += 10
            else:
                model["dis"] += 5
        E2_comment("dis", var+str(i))


def coord_var_dis3(var, num, From, To, sub=False):
    wk.book = "koma_sim3_py.xlsx"
    X = [3, 9]
    for k in X:
        for i in num:
            model = model_size()
            model["x"] = k
            model[var] = i
            model["dis"] = From
            E2_org("dis")
            while model["dis"] <= To:
                E2main(Run3(model), model, "dis")
                if sub == True:
                    model["dis"] += 10
                else:
                    model["dis"] += 5
            E2_comment("dis", var+str(i))
            E2_comment("dis", "x"+str(k)+"mm")
    wk.book = "koma_sim2_py.xlsx"


def coord_var_dis(var, num, From, To, sub=False):
    X = [3, 9]
    for k in X:
        for i in num:
            model = model_size()
            model["x"] = k
            model[var] = i
            model["dis"] = From
            E2_org("dis")
            while model["dis"] <= To:
                E2main(Run(model), model, "dis")
                if sub == True:
                    model["dis"] += 10
                else:
                    model["dis"] += 5
            E2_comment("dis", var+str(i))
            E2_comment("dis", "x"+str(k)+"mm")


def coord_var(var, From, To, X_one=False):
    model = model_size()
    if X_one != False:
        E2_org(var)
        model[var] = From
        model["x"] = X_one
        if var == "mag_num" and From >= 7:
            model["mesh"] = 0.6
        while model[var] <= To:
            E2main(Run(model), model, var)
            if var == "mag_num":
                model[var] += 1
            else:
                model[var] += 5
        E2_comment(var, "x"+str(X_one)+","+var+str(From))
    else:
        X = 3
        while X <= 15:
            model["x"] = X
            E2_org(var)
            Var = From
            while Var <= To:
                model[var] = Var
                E2main(Run(model), model, var)
                if var != "mag_num":
                    Var += 5
                else:
                    Var += 1
                    if Var >= 7:
                        model["mesh"] = 0.6
                E2_comment(var, "x"+str(X)+"mm")
            X += 3


def var_only(var, From, To):
    model = model_size()
    Var = From
    E2_org(var)
    while Var <= To:
        model[var] = Var
        E2main(Run(model), model, var)
        if var != "mag_num":
            Var += 5
        else:
            Var += 1
            if Var >= 7:
                model["mesh"] = 0.6
    E2_comment(var, var+str(From)+"~"+str(To))
