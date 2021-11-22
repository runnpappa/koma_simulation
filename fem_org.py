from fem_excel import *
from fem_graph import *

E3_heatmap_move()
E3_heatmap_xz()
Data_z=E3_heatmap_data("z")
Data_x=E3_heatmap_data("x")
# GraphMain3(Data_z)
GraphMain3(Data_x)