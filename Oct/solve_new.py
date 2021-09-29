import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import scipy
from scipy.integrate import simps

csvname = 5387
#print(csvname)
## import data into py
array_0 = pd.read_csv(f'I:/test/{csvname}-1.csv',header=None,skiprows=45,nrows=4000)#draw data from .csv
array_0 = array_0.values # dataframe --> array
array_1 = pd.read_csv(f'I:/test/{csvname}-2.csv',header=None,skiprows=545,nrows=4501)
array_1 = array_1.values # dataframe --> array
X_trans=1; #the voltage/1V
array_2 = pd.read_csv(f'I:/test/{csvname}-3.csv',header=None,skiprows=135,nrows=4911)
array_2 = array_2.values # dataframe --> array
array_2_1 = np.zeros((2455,2))
for x in range(2455):
    array_2_1[x,] = array_2[2*x,]
print(len(array_2))
'''
array_4 = pd.read_csv(f'I:/test/{csvname}.csv',header=None,skiprows=2846,nrows=2200)
array_4 = array_4.values # dataframe --> array
#print(array_1.shape) print(array_0.shape)

## Define function
def function_50(x): #process data so that 5M and 50M can be connected
    return x-30
def function_60(x): #process data so that 5M and 50M can be connected
    return x-40
array_1[:,1]=function_50(array_1[:,1]) #replace the original data
array_2[:,1]=function_50(array_2[:,1]) #replace the original data
array_4[:,1]=function_60(array_4[:,1]) #replace the original data
array_0 = np.vstack((array_0,array_1,array_2,array_4)) #connect data

def function_T(x,x_trans):
    y = 0.001*50*10**((x-10)/10)/((x_trans*4.2*10**(-8))**2) # VNPSD-->FNPSD
    return y

## connect arraies
array_3 = np.zeros((len(array_0),2))
array_3[:,1]=function_T(array_0[:,1], X_trans)
array_3[:,0]=array_0[:,0]


## get the point about beta-line
for x in range(9500):
    if array_3[x,1] >= 8*np.log(2)*array_3[x,0]/((np.pi)**2):
        # determine if the beta line crosses the PSD
        dot = x
    else:
        break

## integrate the area
integrals = []
for i in range(dot):
    integrals.append(scipy.integrate.
                     trapz(array_3[:i+1,1],array_3[:i+1,0]))
#for i in integrals:
 #   print(i)
linewidth = np.sqrt(8*np.log(2)*integrals[dot-1])
print(linewidth)# print linewidth


## plot FNPSD vs Freq
plt.xlim(1e3, 1e9)
plt.ylim(1,1e7)
plt.xlabel("Frequency (Hz)")
plt.ylabel("FNPSD")
plt.loglog(array_3[:,0],array_3[:,1]) #plot with double log
#plt.text(1e5,10,round(linewidth), fontsize=15)#add text to show linewidth 
plt.show()
'''