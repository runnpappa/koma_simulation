from fem_Main_2 import FemtetMain as Run
from fem_excel import *
from fem_MagSize import model_size
from fem_Run2Main import *
from fem_Main_3 import FemtetMain as Run3

model = model_size()

model["dis"] = 30
model["x"] = 5

while model["dis"] <= 60:
    E2main(Run(model), model, "dis")
    model["dis"] += 5
