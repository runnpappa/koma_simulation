from fem_Main_2 import FemtetMain as Run
from fem_excel import *
from fem_MagSize import model_size
from fem_Run2Main import *
from fem_Main_3 import FemtetMain as Run3

var_sub3(36, 5, 75, 100)

list = [[30, 4, 70], [32, 7, 70], [34, 6, 50], [35, 3, 40],
        [35, 9, 40], [37, 7, 65], [38, 9, 70], [39, 10, 50]]

for i in list:
    var_sub3(i[0], i[1], i[2]-1)
