from fem_excel import *
from fem_graph import *

E3_heatmap_move("data_z")
GraphMain3(E3_heatmap_data("heatmap_z"))
E3_heatmap_move("data_x")
GraphMain3(E3_heatmap_data("heatmap_x"))
E3_heatmap_xz()
GraphMain3(E3_heatmap_data("heatmap_xz"))