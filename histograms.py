from tkinter.tix import NoteBook
import matplotlib
import plotly 
import pandas as pd 
import json
import numpy as np
from plotly.offline import download_plotlyjs, init_notebook_mode, plot, iplot
init_notebook_mode(connected=True)
import matplotlib.pyplot as plt
#import cufflinks as cf
#cf.go_offline()


print(plotly.__version__)
cols=[0,2,3,4,5,6,7,8]
rows=[1,2,3,4,6,8,10,12]
years=""
list_for_years=""
i=0
j=0
standart = "55555"
titles_for_histogram=""
excel_data=pd.read_excel('.vscode\gitwork\histograms\excel.xlsx')
#print('excel_data=',excel_data)
data=pd.DataFrame(excel_data)
print("Content from excel file:\n", data)
#modified_data=pandas.read_excel('.vscode\gitwork\histograms\excel.xlsx',usecols=cols,useindex=rows)
#print('Mod Data:',modified_data)
print("Data from DataFrame\n",data.to_string())
data.items()
for col_name, data in data.items():
    col_name_in_str=str(col_name)
    if (col_name_in_str!='NoN' and col_name_in_str < standart ):
        col_name_in_str=str(col_name)
        years+=col_name_in_str
len_years=len(years)
for j in range(0,7):
    years_for_j=years[i:i+4]
    #value_for_enter="!".join(years_for_j)
    list_for_years+=(years_for_j + ",")
    i+=4
split_sheet=(list_for_years.split(","))
print(split_sheet)
#table = pd.Series(index="index" , column="column")
#for j in range(0,13): # Цикл для создания заголовков 
#    if(str(data.at["Unnamed: 0",j])>"NaN"):
#        titles_for_histogram+=data.at["Unnamed: 0",j]
#print("titles_for_histogram\n", titles_for_histogram)
print(data.at["Unnamed: 0","ВТП Казани, тыс. руб."])