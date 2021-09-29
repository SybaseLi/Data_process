import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import scipy

array_0 = pd.read_csv('G:/aug/12.csv',nrows=1110)#draw data from .csv
array_0 = array_0.values

plt.xlim(350,400)
plt.plot(array_0)
plt.show()