import pandas as pd
import scipy.io as sio

matname = '54.32'# defination of csv document
data_train = sio.loadmat(f'{matname}.mat') #.py need to set in the floder
#print(data_train) # ouput the key of the dictionary
array_0 = data_train["ScSm\x00"] # draw the array from the dictionary
length = len(array_0)

i=0
t=0
for x in range(length):
    if array_0[x,0]>=2955 and array_0[x,0]<=2965:#using the data around wavelength
        i = i+1
        if i == 1:
            t = x
            
array_new= array_0[t:(t+i),] #repalce the data
my_df = pd.DataFrame(array_new)
my_df.to_csv(f'{matname}.csv', index=False) 