import pandas as pd
import pandas
from typing import Dict, List
from openpyxl.styles import Alignment
from tkinter.tix import NoteBook
from matplotlib.pyplot import title
from matplotlib.pyplot import legend
import plotly.express as px
import numpy as np
from plotly import graph_objs as go
from plotly.offline import download_plotlyjs, init_notebook_mode, plot, iplot


data = {'ВТП Казани, тыс. руб.': [674739900.0, 748722900.0, 773031100.0, 798154610.75, 824094635.6, 850877711.26, 878531236.88], 
'Добавленная стоимость, тыс. руб.': [191057445.0, 221584931.0, 226103515.0, 230716026.71, 235422633.65, 240225255.38, 245125850.59], 
'ДС Обрабатывающие производства, тыс. руб.': [75373868.0, 80169194.0, 76066267.0, 78783598.344, 79958359.172, 79850189.785, 81156818.408],
'ДС Строительство, тыс. руб.': [8376613.0, 9553660.0, 8194075.0, 8888142.345, 9065600.551, 8888954.769, 9131323.671],
'ДС Транспортировка и хранение, тыс. руб.': [26272378.0, 35407329.0, 2987596.0, 22157415.832, 20690880.642, 15470346.997, 19884660.344],
'ДС Оптовая и розничная торговля, тыс. руб.': [19963824.0, 25665497.0, 36181429.0, 27718999.183, 30453953.188, 32137036.174, 30691239.965],
'ДС Деятельность в области информатизации и связи, тыс. руб.': [6646511.0, 9461344.0, 10263705.0, 8946153.853, 9756847.568, 9857412.822, 9709393.347]}
data_df = pd.DataFrame(data)
start_year = 2017


def get_axis_x(data_df: pandas.DataFrame,start_year: int) -> list:
    '''
    Функция формирует список со значениями(axis_x) для оси Х(горизонтальной)

    На вход функции подается Data Frame со значениями бюджета и категориями,
    а также год начала отчетного периода

    На выход функция подает готовые значения для оси Х (горизонтальная) в виде отчетного периода  в годах 
    '''
    axis_x= ""
    counter=0
    for key in data_df.keys():
        axis_x += str(start_year+counter) + " "
        counter = counter + 1
    return axis_x


# def make_suitable_dataframe(data_df: pandas.DataFrame, axis_x: list) -> pandas.DataFrame:
#     '''
#     Фукция добавляет данные для оси X(горизонтальной) в Data Frame
#     для использования в функции визуализации графиков  

#     На вход функции подается Data Frame со значениями бюджета и категориями,
#     а также значения для оси Х (горизонтальной) 

#     На выход подается новый Data Frame объект с полным набором данных,
#     необходимым для построения графиков
#     '''

#     #added_dict = {"Период": axis_x.split()}
#     #added_dataframe=pd.DataFrame(added_dict)
#     #data_df_for_histogram=data_df.join(added_dataframe)
#     data_df["Период"] = [axis_x.split()]
#     return data_df


def make_histogram(data_df: pandas.DataFrame, axis_x: list):
    '''
    Функция формирует значения для построения графиков
    и строит графики, сохраняя их с заданным расширением 
    
    На вход функции подается Data Frame объект и значения для оси Х (горизонтальной)
    '''
    
    data_df["Период"] = axis_x.split()  
    for key in data_df.keys():
        if key == 'Период': continue
        fig = go.Figure(go.Bar(data_df, x = axis_x , y = data_df[key]))#, color=["c"]*3+["r"]*4, title=key))
        fig.update_layout(
        xaxis_title_text='Период (*красный- прогнозируемый год)',
        yaxis_title_text= key,
        bargap=0.1,
        template='simple_white') 
        fig.write_html(key + '.html',
               include_plotlyjs='cdn',
               full_html=False,
               include_mathjax='cdn')   


axis_x = get_axis_x(data_df, start_year)
make_histogram(data_df,axis_x)