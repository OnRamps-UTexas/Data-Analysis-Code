import pandas as pd
import numpy as np

########## Specify the name of excel with the data ##########
fileData=pd.read_excel('Sample DataSet.xlsx')

########## Removes last 2 characters ##########
for loop in range(0,fileData['Section'].count()):
    temp_String=str(fileData['Section'].iloc[loop])
    temp_String=temp_String[:-2]
    fileData['Section'].iloc[loop]=temp_String

########## Add the extra Row here #########
title_list=[['Points Possible','','','10','10','10','10','10','10']]

########## Add the column headers here ##########
title_DataFrame=pd.DataFrame(title_list, columns=['Student','Course','Section','Assignment 1 Marks',
'Assignment 2 Marks','Assignment 3 Marks','Assignment 4 Marks','Assignment 5 Marks','Total'])


########## Specify the name of the column based on which the spilt needs to be performed ##########
for section in fileData['College_Section_Name'].unique():
    temp_DataFrame=(fileData[ fileData['College_Section_Name']==section])

    Count=np.count_nonzero(fileData.where(fileData['College_Section_Name']==section).dropna()['Eligibility_Status_Final_Simple'])
    
########## Specify the column which contains eligibility status ##########
    E_count=np.count_nonzero(temp_DataFrame.where(temp_DataFrame['Eligibility_Status_Final_Simple']=='Eligible').dropna()['Eligibility_Status_Final_Simple'])
   


    df_length=len(temp_DataFrame['Student'])
    
    temp_DataFrame.loc[df_length]=['Eligibility Count : ',E_count,'','','','','','','','','','']
    
    Filename=str(temp_DataFrame['College_Section_Name'].iloc[0])+'.xlsx'

    writer = pd.ExcelWriter(Filename, engine='xlsxwriter')
    temp_DataFrame.to_excel(writer,'Sheet1')
    writer.save()

    print('File '+Filename+' created succesfully')

