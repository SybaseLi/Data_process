import pandas as pd
import numpy as np
#import matplotlib.pyplot as plt
#import scipy
#from scipy.integrate import simps

array_1 = pd.read_csv('F:/Data/Linewidth_reducation/25/25-basic/40-50.csv',skiprows=1045,nrows=4000)
array_1 = array_1.values # dataframe --> array
X_trans = 1.183

def function_50(x): #process data so that 5M and 50M can be connected
    return x-30
array_1[:,1]=function_50(array_1[:,1]) #replace the original data


def function_T(x,x_trans):
    y = 0.001*50*10**((x-10)/10)/((x_trans*4.2*10**(-8))**2) # VNPSD-->FNPSD
    return y

array_2 = np.zeros((4000,2))
array_2[:,1]=function_T(array_1[:,1], X_trans)
array_2[:,0]=array_1[:,0]

A = np.sum(array_2[:,1],axis=0)
In_line = np.pi*(A/4000)
print(In_line)