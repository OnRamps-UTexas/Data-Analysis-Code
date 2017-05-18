import pandas as pd
import numpy as np
fileData=pd.read_excel('Sample DataSet.xlsx')

fileData.rename(columns={'University' : 'Teacher'}, inplace=True)

#for loop in range(0,fileData['Teacher'].count()):
#    temp_String=str(fileData['Teacher'].iloc[loop])
#    temp_String=temp_String[:-2]
#    fileData['Teacher'].iloc[loop]=temp_String


counter=0
class_mean=np.sum(fileData['Total Marks'])/np.count_nonzero(fileData['Total Marks'])
teachers_Dataframe = pd.DataFrame(columns=['Teacher', 'Average','Submission Rate','Deviation from class mean'])
teachers_Dataframe_temp = pd.DataFrame(columns=['Teacher', 'Average','Submission Rate','Deviation from class mean'])


for teacher in fileData['Teacher'].unique():
    Sum=np.sum(fileData.where(fileData['Teacher']==teacher).dropna()['Total Marks'])
    Count=np.count_nonzero(fileData.where(fileData['Teacher']==teacher).dropna()['Total Marks'])
    if Count==0:
        average=0
    else:
        average=Sum/Count

    total_count=fileData[ fileData['Teacher']==teacher]['Total Marks'].count()
    submission_Rate=Count*100/total_count

    teachers_Dataframe_temp.loc[counter]=[teacher,average,submission_Rate,average-class_mean];
    counter=counter+1
     
teachers_Dataframe=pd.concat([teachers_Dataframe,teachers_Dataframe_temp])
result=teachers_Dataframe.sort(['Average'], ascending=0)
print(result)


writer = pd.ExcelWriter("Teachers Report.xlsx", engine='xlsxwriter')
result.to_excel(writer,'Sheet1')
writer.save()



