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


########## Specify the name of the column based on which the split needs to be performed ##########
for section in fileData['Section'].unique():
    temp_DataFrame=(fileData[fileData['Section']==section])

########## Columns Whose average you want to calculate ##########
    Sum1=np.sum(fileData.where(fileData['Section']==section)['Assignment 1 Marks'])
    Sum2=np.sum(fileData.where(fileData['Section']==section)['Assignment 2 Marks'])
    Sum3=np.sum(fileData.where(fileData['Section']==section)['Assignment 3 Marks'])
    Sum4=np.sum(fileData.where(fileData['Section']==section)['Assignment 4 Marks'])
    Sum5=np.sum(fileData.where(fileData['Section']==section)['Assignment 5 Marks'])

    Sum=np.sum(fileData.where(fileData['Section']==section)['Total'])


    
    Count=np.count_nonzero(temp_DataFrame.where(temp_DataFrame['Section']==section)['Total'])
    if Count==0:
        average=0
        average1=0
        average2=0
        average3=0
        average4=0
        average5=0

    else:
        average=Sum/Count
        average1=Sum1/Count
        average2=Sum2/Count
        average3=Sum3/Count
        average4=Sum4/Count
        average5=Sum5/Count
 


########## Specify the name of the column based on which the sort needs to be performed ##########
    result=temp_DataFrame.sort(['Total'], ascending=0)
    result=pd.concat([title_DataFrame,result])

    df_length=len(result['Student'])

########## Add spaces for columns with average ##########    
    result.loc[df_length]=['Average','','',average1,average2,average3,average4,average5,average]
    

    Filename=str(temp_DataFrame['Section'].iloc[0])+'.xlsx'
    writer = pd.ExcelWriter(Filename, engine='xlsxwriter')
    result.to_excel(writer,'Sheet1')
    writer.save()

    print('File '+Filename+' created succesfully')


