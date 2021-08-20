import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import scipy
from scipy.integrate import simps

## import data into py
array_0 = pd.read_csv('F:/Data/Linewidth——reducation/8-18/25/10-5.csv',skiprows=45,nrows=5000)#draw data from .csv
array_0 = array_0.values # dataframe --> array
array_1 = pd.read_csv('F:/Data/Linewidth——reducation/8-18/25/10-50.csv',skiprows=545,nrows=4500)
array_1 = array_1.values # dataframe --> array
X_trans=0.865; #the voltage/1V
#print(array_1.shape) print(array_0.shape)

## Define function
def function_50(x): #process data so that 5M and 50M can be connected
    return x-30
array_1[:,1]=function_50(array_1[:,1]) #replace the original data
array_0 = np.vstack((array_0,array_1)) #connect data

def function_T(x,x_trans):
    y = 0.001*50*10**((x-10)/10)/((x_trans*4.2*10**(-8))**2) # VNPSD-->FNPSD
    return y

## connect arraies
array_2 = np.zeros((9500,2))
array_2[:,1]=function_T(array_0[:,1], X_trans)
array_2[:,0]=array_0[:,0]

## get the point about beta-line
for x in range(9500):
    if array_2[x,1] >= 8*np.log(2)*array_2[x,0]/((np.pi)**2):
        # determine if the beta line crosses the PSD
        dot = x
    else:
        break

## integrate the area
integrals = []
for i in range(dot):
    integrals.append(scipy.integrate.
                     trapz(array_2[:i+1,1],array_2[:i+1,0]))
#for i in integrals:
 #   print(i)
linewidth = np.sqrt(8*np.log(2)*integrals[dot-1])
#print(linewidth)# print linewidth

## plot FNPSD vs Freq
plt.xlim(1e3, 5e7)
plt.ylim(1,1e7)
plt.xlabel("Frequency (Hz)")
plt.ylabel("FNPSD")
plt.loglog(array_2[:,0],array_2[:,1]) #plot with double log
plt.text(1e5,10,round(linewidth), fontsize=15)#add text to show linewidth 
plt.show()