"""
step 1 import the required librries and make sures all nescessary libraies are
installed in your system"""

import shutil # Used for copying files
import os
import pandas as pd  # Used for handling Excel files


##########
##extract and save the signals in data frame using pandas

##to do in future make it as input
signaldataframe=pd.read_excel('Signals.xlsx')

headerlist =[signaldataframe.columns.values.tolist()]##the labels

datalist = signaldataframe.values.tolist()

####### main function
##it take two text path 
def generateloops(destination,template_F):

    #make sure the dirctory exist
    os.makedirs(destination, exist_ok=True)
    os.makedirs(template_F, exist_ok=True)
    os.chdir(template_F)


    for signal in datalist:
        templeteName = signal[0]
        root, extension = os.path.splitext(templeteName)

        #full destination path ex:: d\mahmou\contact1.dwg
        newname= signal [1]+ extension#contact1.dwg

        destination_path = os.path.join(destination,newname)##d\mahmou\contact1.dwg

        #copy the template and rename it for every signal in the excel
        shutil.copy2(signal[0], destination_path)

#get the destination and the template folders as inputs 
destination = input("Enter your destination folder path: ")
print("\n")
template_F = input("Enter your templates folder path: ")

#####################
generateloops(destination,template_F)
#####################
done= input("\n done")
        
        

        
   
