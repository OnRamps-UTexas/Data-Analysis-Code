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


########## Specify the name of the column based on which the sort needs to be performed ##########
    result=temp_DataFrame.sort(['Total'], ascending=0)
    result=pd.concat([title_DataFrame,result])

    Filename=str(temp_DataFrame['Section'].iloc[0])+'.xlsx'
    writer = pd.ExcelWriter(Filename, engine='xlsxwriter')
    result.to_excel(writer,'Sheet1')
    writer.save()

    print('File '+Filename+' created succesfully')


