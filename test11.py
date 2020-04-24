import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
import matplotlib.gridspec as gridspec
from matplotlib import animation
import pandas as pd

df = pd.DataFrame({'a':[1,2,3],'b':[4,5,6]})
# df.set_index('a',inplace=True)
df = df.reindex(list(range(1,32)),fill_value=np.nan)
print(df)