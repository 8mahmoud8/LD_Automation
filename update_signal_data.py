import shutil
import xlwt #to work with xls old excel file
import os
import pandas as pd
##########################
#get the signals data and save it

signaldataframe=pd.read_excel('Signals.xlsx')

headerlist =signaldataframe.columns.values.tolist()

datalist = signaldataframe.values.tolist()
##########################################
###get the file full pat  as input
Cadfile =input ("enter the full autocad execel path \n for ex D:mahmoud.hamada\LD.XLS \n")
print("\n ")
#get the cad data and save it
Cad_data_frame = pd.read_excel(Cadfile)

Cad_headerlist =Cad_data_frame.columns.values.tolist()

CAd_datalist = Cad_data_frame.values.tolist()

#########################################

#search the cad file and edite it row by row

for i in range(len(Cad_data_frame)):
    ##in cad excel file get the file name cell for the i row
    fileName=Cad_data_frame.loc[i,'(FILENAME)']

     #extract the signal name from the file name 
    signalName, extension = os.path.splitext(fileName)

    ##in cad excel file get the tagname cell 
    tagName=Cad_data_frame.loc[i,'TAGNAME']

    # handling the uper case by make list of the signals in upper cas
    upersignal=list(map(str.upper, signaldataframe['signal_name']))
    
    #check the signal is exsited in the signal execl data
    if(signalName in upersignal):
        #get the signal index in signal excel
        signal_index=[upersignal.index(signalName)]

        ##check the tagname is existed in signal excel data
        if (tagName in headerlist):

            updated_description= signaldataframe.loc[signal_index,tagName]

            Cad_data_frame.loc[i,'DESC2']=updated_description.iloc[0]
            
            

        
######################
#################write the updated data frame in xls file 
#############the xls file is old so i did it in old way with xlwt librray 
##################

#create a workbook an 3 sheet as the autocad execel file has 2 empty sheets

wb = xlwt.Workbook()
ws = wb.add_sheet('Sheet1')
ws2 = wb.add_sheet('Sheet2')
ws3 = wb.add_sheet('Sheet3')


#########update the data and headers lists 

Cad_headerlist =Cad_data_frame.columns.values.tolist()
CAd_datalist = Cad_data_frame.values.tolist()
data = [Cad_headerlist] + CAd_datalist

################


### Write data row by row
for row_index, row in enumerate(data):
    for col_index, value in enumerate(row):
        
        # Clean the value: skip if NaN or empty string
        if pd.isna(value) or value == '':
            continue  # leave the cell empty
        ws.write(row_index, col_index, str(value))  # force to string if needed


# Save the workbook
wb.save(Cadfile)
########
done=input ("done")


    
