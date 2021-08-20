import pandas as pd
import matplotlib.pyplot as plt

## import data into py
array_1 = pd.read_csv('F:/Data/Linewidth——reducation/8-18/25/30-1g.csv',skiprows=45,nrows=5001)#draw data from .csv
array_1 = array_1.values # dataframe --> array

## Define function
#def function_50(x): #process data so that 5M and 50M can be connected
#    return x-50
#array_1[:,1]=function_50(array_1[:,1]) #replace the original data

#plt.xlim(1000, 1000000000)
#plt.ylim(-90,-10)
#plt.xlabel("Frequency (GHz)")
#plt.ylabel("dBm")
#plt.plot(array_1[:,0],array_1[:,1]) #plot with double log
#plt.show()